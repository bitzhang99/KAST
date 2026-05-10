from __future__ import annotations

import configparser
import difflib
import json
import sys
import traceback
from datetime import datetime
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter, column_index_from_string

# 修订记录（按时间顺序）
# V1: 基础版，支持固定路径读取模板与输入文件并写出结果
# V2: 增加字段映射与“接收日期”双写入（创建时间、回复时间/ITM回函时间）
# V3: 增加按“工单编号/ITM编号”去重与模板列对齐
# V4: 增加客户类型识别、客户名称前缀清理、回复人映射
# V5: 增加日期列解析与 Excel 日期格式/列宽设置
# V6: 增加 ini 配置化输入（TEMPLATE_PATH、INPUT_PATHS）并支持多文件
# V7: 增强容错（相似文件名匹配、缺失输入跳过）并补全回复人映射规则
# V8: 固定配置文件名为 merge_requirement_to_template.ini，
#     并将标题映射/回复人映射统一迁移到 ini（Title_Map/TaskAssigner_Map）
# V9: 增加运行日志文件输出（新增数、去重后总数、重复工单编号清单）
# V10: 增加路径存在性预检查与友好报错建议，日志同时写入文件与标准输出
# V11: 新增数据按“创建时间”排序后追加到模板末尾，避免与原始模板数据混编
# V12: 输出 Excel 打开时默认定位到“原模板数据倒数5行”位置，便于续看历史数据
# V13: 读取 ini 使用 utf-8-sig，兼容 Windows 记事本保存的带 BOM 的 UTF-8
# 解决了pyinstaller 打包以后找不到配置文件的问题
# V14：20260408-给“工单状态”默认赋值“进行中”，“事项来源”默认赋值“ITM”
# V15：20260510-以导入“要求编号”映射后的键为唯一键；与模板“业务要求编号”或“工单编号/ITM编号”匹配则整行按导入字段更新；
#      多文件导入先去重；AC/AD/AE 列按模板首行公式模式向下翻译写入；转换失败日志记录到具体记录键与字段；
#      修复重复列名/混合类型排序等导致的 could not convert string to float 类错误
CONFIG_PATH = Path(__file__).with_name("merge_requirement_to_template.ini")
CONFIG_SECTION = "PATHS"

TEMPLATE_SHEET = "模板页"
# [V15] 与模板中行公式同步：按列将首行数据区的公式模式复制到所有数据行（AC/AD/AE）
FORMULA_COLUMN_LETTERS = ("AC", "AD", "AE")
FORMULA_COL_INDEX: tuple[int, ...] = tuple(column_index_from_string(c) for c in FORMULA_COLUMN_LETTERS)


def normalize(s: object) -> str:
    if s is None or pd.isna(s):
        return ""
    return str(s).strip()


def build_output_path(template_path: Path) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return template_path.with_name(f"{template_path.stem}_{ts}{template_path.suffix}")


def build_log_path(script_path: Path) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return script_path.with_name(f"{script_path.stem}_{ts}.log")


def parse_input_paths(raw_value: str) -> list[Path]:
    """
    INPUT_PATHS supports:
    - one path per line
    - comma-separated paths
    - mixed line/comma style
    """
    normalized = raw_value.replace("\r\n", "\n").replace("\r", "\n")
    normalized = normalized.replace("，", ",").replace("；", ";")
    parts: list[str] = []
    for line in normalized.split("\n"):
        semicolon_split = []
        for block in line.split(";"):
            semicolon_split.extend(block.split(","))
        for piece in semicolon_split:
            candidate = piece.strip().strip('"').strip("'")
            if candidate:
                parts.append(candidate)
    return [Path(p) for p in parts]


def normalize_filename_key(name: str) -> str:
    return (
        name.strip()
        .replace("（", "(")
        .replace("）", ")")
        .replace("【", "[")
        .replace("】", "]")
        .replace("，", ",")
        .replace("。", ".")
        .replace("：", ":")
        .replace(" ", "")
        .lower()
    )


def resolve_similar_excel_path(path: Path) -> Path:
    if path.exists():
        return path
    parent = path.parent
    if not parent.exists():
        return path

    target_key = normalize_filename_key(path.name)
    candidates = list(parent.glob("*.xlsx")) + list(parent.glob("*.xlsm"))
    for candidate in candidates:
        if normalize_filename_key(candidate.name) == target_key:
            return candidate
    return path


def build_path_hint(path: Path) -> str:
    parent = path.parent
    if not parent.exists():
        return (
            f"路径不存在: {path}\n"
            f"建议: 先确认目录是否存在 -> {parent}\n"
            "建议: 检查是否把中文括号/英文括号写错，或是否缺少盘符前缀。"
        )

    excel_candidates = [p.name for p in parent.iterdir() if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm", ".xls"}]
    if not excel_candidates:
        return (
            f"路径不存在: {path}\n"
            f"建议: 目录存在但未发现 Excel 文件 -> {parent}\n"
            "建议: 检查文件是否尚未保存、后缀是否正确。"
        )

    matches = difflib.get_close_matches(path.name, excel_candidates, n=3, cutoff=0.45)
    if matches:
        candidates_text = "\n".join(f"  - {name}" for name in matches)
        return (
            f"路径不存在: {path}\n"
            "建议: 发现同目录相似文件名，可能是括号/标点/全半角差异：\n"
            f"{candidates_text}"
        )

    sample = "\n".join(f"  - {name}" for name in excel_candidates[:5])
    return (
        f"路径不存在: {path}\n"
        "建议: 同目录存在以下 Excel 文件，请核对是否应使用其中之一：\n"
        f"{sample}"
    )


def parse_title_map(raw_value: str) -> dict[str, list[str]]:
    if not raw_value.strip():
        raise ValueError("配置项 Title_Map 不能为空")
    try:
        parsed = json.loads(raw_value)
    except json.JSONDecodeError as exc:
        raise ValueError(f"Title_Map 不是合法 JSON: {exc}") from exc

    if not isinstance(parsed, dict):
        raise ValueError("Title_Map 必须是 JSON 对象（键值映射）")

    normalized_map: dict[str, list[str]] = {}
    for src_col, target in parsed.items():
        src_name = normalize(src_col)
        if src_name == "":
            continue
        if isinstance(target, str):
            target_names = [normalize(target)]
        elif isinstance(target, list):
            target_names = [normalize(x) for x in target if isinstance(x, str)]
        else:
            raise ValueError(f"Title_Map 中字段 {src_col} 的映射值必须是字符串或字符串数组")

        target_names = [x for x in target_names if x != ""]
        if not target_names:
            continue
        normalized_map[src_name] = target_names

    if not normalized_map:
        raise ValueError("Title_Map 解析后为空，请检查配置")
    return normalized_map


def parse_task_assigner_map(raw_value: str) -> dict[str, str]:
    if not raw_value.strip():
        raise ValueError("配置项 TaskAssigner_Map 不能为空")
    try:
        parsed = json.loads(raw_value)
    except json.JSONDecodeError as exc:
        raise ValueError(f"TaskAssigner_Map 不是合法 JSON: {exc}") from exc

    if not isinstance(parsed, dict):
        raise ValueError("TaskAssigner_Map 必须是 JSON 对象（键值映射）")

    normalized_map: dict[str, str] = {}
    for org_name, assignee in parsed.items():
        org_key = normalize(org_name)
        assignee_name = normalize(assignee)
        if org_key and assignee_name:
            normalized_map[org_key] = assignee_name

    if not normalized_map:
        raise ValueError("TaskAssigner_Map 解析后为空，请检查配置")
    return normalized_map


def load_config_from_ini(
    config_path: Path,
) -> tuple[Path, list[Path], list[Path], dict[str, list[str]], dict[str, str]]:
    if not config_path.exists():
        raise FileNotFoundError(
            f"未找到配置文件: {config_path}\n"
            "请在脚本同目录创建 merge_requirement_to_template.ini，"
            "并在 [PATHS] 中配置 TEMPLATE_PATH、INPUT_PATHS、Title_Map、TaskAssigner_Map。"
        )

    parser = configparser.ConfigParser()
    parser.read(config_path, encoding="utf-8-sig")

    if CONFIG_SECTION not in parser:
        raise ValueError(f"配置文件缺少节: [{CONFIG_SECTION}]")

    section = parser[CONFIG_SECTION]
    template_raw = section.get("TEMPLATE_PATH", "").strip()
    inputs_raw = section.get("INPUT_PATHS", "").strip()
    title_map_raw = section.get("Title_Map", "").strip()
    task_assigner_map_raw = section.get("TaskAssigner_Map", "").strip()

    if not template_raw:
        raise ValueError("配置项 TEMPLATE_PATH 不能为空")
    if not inputs_raw:
        raise ValueError("配置项 INPUT_PATHS 不能为空")

    template_path = resolve_similar_excel_path(Path(template_raw))
    input_paths = [resolve_similar_excel_path(p) for p in parse_input_paths(inputs_raw)]
    if not input_paths:
        raise ValueError("INPUT_PATHS 未解析到有效文件路径")

    if not template_path.exists():
        raise FileNotFoundError(
            "TEMPLATE_PATH 文件不存在，请检查 [PATHS] 中 TEMPLATE_PATH 配置。\n"
            + build_path_hint(template_path)
        )
    missing_inputs = [p for p in input_paths if not p.exists()]
    existing_inputs = [p for p in input_paths if p.exists()]
    if not existing_inputs:
        hints = "\n\n".join(build_path_hint(p) for p in input_paths)
        raise FileNotFoundError(
            "INPUT_PATHS 中没有可用文件，请检查 [PATHS] 中 INPUT_PATHS 配置。\n"
            f"{hints}"
        )

    title_map = parse_title_map(title_map_raw)
    task_assigner_map = parse_task_assigner_map(task_assigner_map_raw)
    return template_path, existing_inputs, missing_inputs, title_map, task_assigner_map


def read_input_with_fallback(path: Path) -> pd.DataFrame:
    """
    Read preferred sheet 'Sheet0' when present, otherwise fallback to first sheet.
    """
    xls = pd.ExcelFile(path)
    preferred = "Sheet0"
    sheet_name = preferred if preferred in xls.sheet_names else xls.sheet_names[0]
    return pd.read_excel(path, sheet_name=sheet_name)


def infer_customer_type(customer_name: object) -> object:
    text = normalize(customer_name)
    if text == "":
        return pd.NA
    has_hq = "对公客户-总行级客户-" in text
    has_branch = "对公客户-分行级重点客户-" in text
    if has_hq:
        return "总行级客户"
    if has_branch:
        return "分行级重点客户"
    return pd.NA


def clean_customer_name(customer_name: object) -> object:
    text = normalize(customer_name)
    if text == "":
        return pd.NA
    text = text.replace("对公客户-总行级客户-", "")
    text = text.replace("对公客户-分行级重点客户-", "")
    return text.strip() or pd.NA

#  该程序将 总行部分的需求映射到李卫和张峰来处理
def infer_responder(org_name: object, responder_mapping: dict[str, str]) -> object:
    text = normalize(org_name)
    if text == "":
        return pd.NA
    if "中国" in text or "总行" in text:
        return "李卫、张峰"
    return responder_mapping.get(text, pd.NA)

# 如下代码设置日期格式字段为日期  步骤1
def parse_datetime_cell(value: object) -> object:
    if value is None or pd.isna(value):
        return pd.NA
    ts = pd.to_datetime(value, errors="coerce")
    if pd.isna(ts):
        return pd.NA
    return ts.to_pydatetime()


# [V15] 合并重复表头（归一化后同名）时保留首列，避免 select 出 DataFrame 导致 to_datetime/sort 异常
def consolidate_duplicate_header_columns(df: pd.DataFrame, log_lines: list[str], label: str) -> pd.DataFrame:
    norm_cols = [normalize(str(c)) for c in df.columns]
    seen: set[str] = set()
    drop_list: list[int | str] = []
    for i, c in enumerate(df.columns):
        key = norm_cols[i] or f"_col{i}"
        if key in seen:
            log_lines.append(f"[列名警告][{label}] 归一化后重复的列已忽略: 原始列名={c!r} (键={key})")
            drop_list.append(c)
        else:
            seen.add(key)
    out = df.drop(columns=drop_list, errors="ignore").copy()
    out.columns = [normalize(str(c)) for c in out.columns]
    return out


# [V15] 模板行匹配键：优先“业务要求编号”，否则“工单编号/ITM编号”（与导入侧“要求编号”映射键对齐）
def template_match_key_series(df: pd.DataFrame) -> pd.Series:
    has_biz = "业务要求编号" in df.columns
    has_ticket = "工单编号/ITM编号" in df.columns
    if not has_biz and not has_ticket:
        return pd.Series([""] * len(df), index=df.index)
    k_biz = df["业务要求编号"].map(normalize) if has_biz else pd.Series("", index=df.index)
    k_ticket = df["工单编号/ITM编号"].map(normalize) if has_ticket else pd.Series("", index=df.index)
    merged = k_biz.where((k_biz != "") & (~k_biz.isna()), k_ticket)
    merged = merged.fillna("").map(lambda x: normalize(x))
    return merged


# [V15] 在删除数据行前抓取 AC/AD/AE 上首个含公式的行，作为向下复制的锚点
def capture_formula_templates_for_columns(
    ws, col_indexes: tuple[int, ...], scan_rows: int = 30
) -> dict[int, tuple[str, str]]:
    specs: dict[int, tuple[str, str]] = {}
    max_scan = min(ws.max_row, scan_rows)
    for r in range(2, max_scan + 1):
        row_has_formula = False
        for c in col_indexes:
            cell = ws.cell(row=r, column=c)
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                col_letter = get_column_letter(c)
                specs[c] = (v, f"{col_letter}{r}")
                row_has_formula = True
        if row_has_formula:
            break
    return specs


# [V15] 将锚点公式按目标行翻译写入（列模式复制）
def apply_formula_templates_to_rows(
    ws, specs: dict[int, tuple[str, str]], row_from: int, row_to: int, log_lines: list[str]
) -> None:
    if row_to < row_from or not specs:
        return
    for r in range(row_from, row_to + 1):
        for c, (formula, origin) in specs.items():
            col_letter = get_column_letter(c)
            dest = f"{col_letter}{r}"
            try:
                ws.cell(row=r, column=c).value = Translator(formula, origin=origin).translate_formula(dest)
            except Exception as exc:
                log_lines.append(f"[公式写入警告] 行={r} 列={col_letter} 原因={exc}")


# [V15] 导入行覆盖模板行：导入非空则更新，否则保留模板原值
def merge_template_row_with_import(base: pd.Series, imp: pd.Series, headers: list[str]) -> pd.Series:
    out = base.reindex(headers).copy()
    for col in headers:
        if col not in imp.index:
            continue
        v = imp[col]
        if pd.isna(v):
            continue
        out[col] = v
    return out


# [V15] 日期列逐格转换，失败时记录「记录键 + 字段 + 原始值」
def parse_datetime_series_logged(
    series: pd.Series, row_keys: pd.Series, col_name: str, log_lines: list[str]
) -> pd.Series:
    out: list[object] = []
    for i in range(len(series)):
        raw = series.iloc[i]
        rk = normalize(row_keys.iloc[i]) if i < len(row_keys) else ""
        try:
            out.append(parse_datetime_cell(raw))
        except Exception as exc:
            # [V15] 捕获含 could not convert string to float 在内的各类转换异常并落日志
            log_lines.append(
                f"[字段转换失败] 记录键(工单编号/ITM编号)={rk or '(空)'} 字段={col_name} "
                f"原始值={raw!r} 错误类型={type(exc).__name__} 错误信息={exc}"
            )
            out.append(pd.NA)
    return pd.Series(out, index=series.index)


# [V15] 对「创建时间」排序列安全 to_datetime，避免重复列返回 DataFrame 或混合类型触发 float 转换异常
def safe_series_to_datetime_for_sort(s: pd.Series, log_lines: list[str], label: str) -> pd.Series:
    if isinstance(s, pd.DataFrame):
        log_lines.append(f"[排序警告][{label}] 「创建时间」对应多列，已仅使用第一列参与排序")
        s = s.iloc[:, 0]
    try:
        return pd.to_datetime(s, errors="coerce")
    except (ValueError, TypeError) as exc:
        log_lines.append(f"[排序警告][{label}] 「创建时间」to_datetime 失败({exc})，已回退为不作为时间排序")
        return pd.Series([pd.NaT] * len(s), index=s.index)


def collect_duplicate_ids(df: pd.DataFrame, unique_key: str) -> list[str]:
    if unique_key not in df.columns:
        return []
    keys = df[unique_key].map(normalize)
    keys = keys[keys != ""]
    key_counts = keys.value_counts()
    duplicates = key_counts[key_counts > 1].index.tolist()
    return [str(x) for x in duplicates]


def emit_log(lines: list[str], log_path: Path) -> None:
    text = "\n".join(lines) + "\n"
    # [V10] 日志同时写入文件与标准输出，便于即时查看运行结果
    print(text, end="")
    with log_path.open("w", encoding="utf-8") as f:
        f.write(text)


def set_initial_view_to_template_tail(ws, original_template_rows: int) -> None:
    # [R12-01] 生成文件时设置初始可视区域：定位到“原模板数据倒数5行”的起始位置
    # 说明：模板数据从第2行开始（第1行为表头），若模板不足5行则回退到第2行。
    if original_template_rows <= 0:
        target_row = 2
    else:
        first_of_last_five_data_row = max(1, original_template_rows - 4)
        target_row = first_of_last_five_data_row + 1  # +1 对齐到 Excel 实际行号（含表头）

    target_cell = f"A{target_row}"
    ws.sheet_view.topLeftCell = target_cell
    if ws.sheet_view.selection:
        ws.sheet_view.selection[0].activeCell = target_cell
        ws.sheet_view.selection[0].sqref = target_cell


def main() -> None:
    # [V6->V8] 从固定配置文件读取路径、多输入文件、标题映射与回复人映射
    # [V15] 全程累积 log_lines，异常时亦可写出已收集的诊断信息
    log_path = build_log_path(Path(__file__))
    log_lines: list[str] = []
    try:
        template_path, input_paths, missing_inputs, title_map, responder_mapping = load_config_from_ini(CONFIG_PATH)
        if missing_inputs:
            print("警告: 以下输入文件不存在，已自动跳过：")
            for p in missing_inputs:
                print(f" - {p}")

        # [V3] 读取模板；[V15] 合并重复表头，避免「创建时间」等列变成 DataFrame 触发 float 转换异常
        template_df = pd.read_excel(template_path, sheet_name=TEMPLATE_SHEET)
        original_template_rows = len(template_df)
        template_df = consolidate_duplicate_header_columns(template_df, log_lines, "模板")
        template_headers = list(template_df.columns)
        template_header_set = set(template_headers)

        merged_parts: list[pd.DataFrame] = []
        for p in input_paths:
            src = read_input_with_fallback(p)
            src = consolidate_duplicate_header_columns(src, log_lines, f"输入:{p.name}")

            out = pd.DataFrame(index=src.index)
            for src_col, target_cols in title_map.items():
                for target_col in target_cols:
                    out[target_col] = src[src_col] if src_col in src.columns else pd.NA

            # V14:20260408 对一些固定的列做默认值填写
            out["事项来源"] = "ITM"
            out["当前状态"] = "进行中"

            for col in template_headers:
                if col not in out.columns:
                    out[col] = pd.NA
            out = out[template_headers]
            out = out[[c for c in out.columns if c in template_header_set]]
            merged_parts.append(out)

        if not merged_parts:
            raise ValueError("没有可合并的输入数据，请检查 INPUT_PATHS 配置。")
        all_new_rows = pd.concat(merged_parts, ignore_index=True)
        import_raw_row_count = len(all_new_rows)

        # [V15] 唯一键：导入侧「要求编号」经 Title_Map 映射后的「工单编号/ITM编号」；与模板「业务要求编号」或「工单编号/ITM编号」匹配则更新该行
        unique_key = "工单编号/ITM编号"
        if unique_key not in template_headers:
            raise ValueError(f"模板页缺少唯一索引列: {unique_key}")

        new_rows = all_new_rows.copy()
        new_rows["_ik"] = new_rows[unique_key].map(normalize)
        new_rows = new_rows[new_rows["_ik"] != ""].reset_index(drop=True)
        import_nonempty_key_count = len(new_rows)
        new_rows = new_rows.drop_duplicates(subset=["_ik"], keep="last").reset_index(drop=True)
        new_rows = new_rows.drop(columns=["_ik"])
        import_dedup_dropped = import_nonempty_key_count - len(new_rows)

        imp_dict: dict[str, pd.Series] = {}
        for _, row in new_rows.iterrows():
            imp_dict[normalize(row[unique_key])] = row

        template_work = template_df.copy()
        template_work["_mk"] = template_match_key_series(template_work)
        template_work = template_work[template_work["_mk"] != ""].reset_index(drop=True)
        template_work = template_work.drop_duplicates(subset=["_mk"], keep="last").reset_index(drop=True)
        template_keys_set = set(template_work["_mk"].map(normalize))

        dup_source = pd.DataFrame(
            {unique_key: pd.concat([template_work["_mk"], new_rows[unique_key].map(normalize)], ignore_index=True)}
        )
        duplicate_ids = collect_duplicate_ids(dup_source, unique_key)

        result_rows: list[pd.Series] = []
        keys_updated: list[str] = []
        for _, trow in template_work.iterrows():
            mk = normalize(str(trow["_mk"]))
            base = trow.drop(labels=["_mk"])
            if mk in imp_dict:
                # [V15] 模板与导入编号一致：用导入非空字段覆盖模板该行
                result_rows.append(merge_template_row_with_import(base, imp_dict[mk], template_headers))
                keys_updated.append(mk)
            else:
                result_rows.append(base.reindex(template_headers))

        append_keys = [k for k in imp_dict if k not in template_keys_set]
        append_df = pd.DataFrame([imp_dict[k] for k in append_keys], columns=template_headers)
        if len(append_df) and "创建时间" in append_df.columns:
            append_df["_st"] = safe_series_to_datetime_for_sort(append_df["创建时间"], log_lines, "追加排序")
            append_df["_sk"] = append_df[unique_key].map(normalize).astype(str)
            try:
                append_df = append_df.sort_values(
                    by=["_st", "_sk"], ascending=[True, True], na_position="last"
                ).drop(columns=["_st", "_sk"])
            except (ValueError, TypeError) as exc:
                log_lines.append(f"[排序失败] {exc}，已回退为仅按「{unique_key}」字符串排序")
                append_df = append_df.drop(columns=["_st", "_sk"], errors="ignore").sort_values(
                    by=[unique_key], key=lambda s: s.astype(str)
                )

        combined = pd.concat(
            [
                pd.DataFrame(result_rows, columns=template_headers),
                append_df.reindex(columns=template_headers),
            ],
            ignore_index=True,
        )

        customer_name_col = "客户名称"
        customer_type_col = "客户类型"
        if customer_name_col in combined.columns:
            if customer_type_col in combined.columns:
                inferred = combined[customer_name_col].map(infer_customer_type)
                existing = combined[customer_type_col]
                combined[customer_type_col] = inferred.where(~inferred.isna(), existing)
            combined[customer_name_col] = combined[customer_name_col].map(clean_customer_name)

        creator_org_col = "创建人机构"
        responder_col = "回复人"
        if creator_org_col in combined.columns and responder_col in combined.columns:
            combined[responder_col] = combined[creator_org_col].map(
                lambda x: infer_responder(x, responder_mapping)
            )

        date_cols = ["创建时间", "回复时间/ITM回函时间"]
        row_keys = combined[unique_key].map(normalize) if unique_key in combined.columns else pd.Series([""] * len(combined))
        for col in date_cols:
            if col in combined.columns:
                # [V15] 日期列逐格记录转换失败（记录键 + 字段名 + 原始值）
                combined[col] = parse_datetime_series_logged(combined[col], row_keys, col, log_lines)

        output_path = build_output_path(template_path)
        wb = load_workbook(template_path)
        ws = wb[TEMPLATE_SHEET]

        # [V15] 在清空数据行前抓取 AC/AD/AE 公式锚点，删除行后再写回并翻译到各行
        formula_specs = capture_formula_templates_for_columns(ws, FORMULA_COL_INDEX)
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)

        start_row = 2
        formula_col_set = set(FORMULA_COL_INDEX)
        for i, row in combined.iterrows():
            excel_r = start_row + i
            for j, col in enumerate(template_headers, start=1):
                if j in formula_col_set:
                    continue
                val = row[col]
                ws.cell(row=excel_r, column=j, value=None if pd.isna(val) else val)

        apply_formula_templates_to_rows(ws, formula_specs, 2, ws.max_row, log_lines)

        date_number_format = "yyyy-mm-dd hh:mm:ss"
        for col in date_cols:
            if col in template_headers:
                col_idx = template_headers.index(col) + 1
                for row_idx in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if cell.value is not None:
                        cell.number_format = date_number_format
                ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = 21

        set_initial_view_to_template_tail(ws, original_template_rows)
        wb.save(output_path)

        log_lines = [
            f"运行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"程序版本修订: V15（要求编号/业务要求编号匹配更新、导入去重、公式列、转换日志）",
            f"模板文件: {template_path}",
            "输入文件:",
            * [f" - {p}" for p in input_paths],
            *(["缺失输入文件(已跳过):", *[f" - {p}" for p in missing_inputs]] if missing_inputs else []),
            f"输出文件: {output_path}",
            f"日志文件: {log_path}",
            f"导入原始记录数(多文件合计): {import_raw_row_count}",
            f"导入侧去重删除行数(多文件重复「{unique_key}」): {import_dedup_dropped}",
            f"与模板匹配并已更新的记录数: {len(keys_updated)}",
            f"纯新增并追加到末尾的记录数: {len(append_df)}",
            f"去重后总记录数: {len(combined)}",
            f"重复键统计用「{unique_key}」(含模板匹配键与导入键合并视角) 重复值个数: {len(duplicate_ids)}",
            f"重复{unique_key}值:",
            * ([f" - {key}" for key in duplicate_ids] if duplicate_ids else [" - 无"]),
        ] + log_lines
        emit_log(log_lines, log_path)
    except Exception:
        log_lines.append("---- 以下为异常发生前/处理过程中的诊断与堆栈 ----")
        log_lines.append(traceback.format_exc())
        emit_log(log_lines, log_path)
        raise


# 解决 pyinstaller 打包后找不到配置文件的问题
def get_base_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    else:
        return Path(__file__).parent

if __name__ == "__main__":
    base_dir = get_base_dir()
    CONFIG_PATH = base_dir / 'merge_requirement_to_template.ini'
    print(f"CONFIG_PATH: {CONFIG_PATH}")

    if not CONFIG_PATH.exists():
        print(f"未找到配置文件: {CONFIG_PATH}")
        sys.exit(1)

    try:
        main()
    except Exception as exc:
        print("程序执行失败，请检查以下错误信息：", file=sys.stderr)
        print(str(exc), file=sys.stderr)
        sys.exit(1)