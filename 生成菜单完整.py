import os
import sys
from collections import defaultdict

import pandas as pd


CONJUNCTIONS = {"&", "and", "+", "with", "to"}
DATA_FILE = "collection.csv"
COLLECTION_DIR = "collection"
TEMPLATE_FILE = "menu(1).csv"
OUTPUT_FILE = "menu_updated.csv"
MAX_DEPTH = 3
MIN_INFERRED_GROUP_WORDS = 2
MIN_INFERRED_GROUP_CHILDREN = 2
CSV_ENCODINGS = ("utf-8-sig", "utf-8", "gb18030")
REQUIRED_TEMPLATE_COLUMNS = (
    "Row #",
    "Menu Item: ID",
    "Menu Item: Parent ID",
    "Menu Item: Title",
    "Menu Item: Resource Handle",
    "Menu Item: URL",
)


def get_application_path():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def normalize_text(value):
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().split())


def read_csv_compat(csv_path):
    last_error = None
    for encoding in CSV_ENCODINGS:
        try:
            return pd.read_csv(
                csv_path,
                encoding=encoding,
                dtype=str,
                keep_default_na=False,
            )
        except UnicodeDecodeError as exc:
            last_error = exc
            continue
    raise ValueError(
        f"读取 CSV 失败：{csv_path}，已尝试编码 {', '.join(CSV_ENCODINGS)}，"
        f"最后错误：{last_error}"
    )


def normalize_title_for_lookup(title):
    return normalize_text(title.replace("&", " & ").replace("+", " + "))


def title_words(title):
    return normalize_title_for_lookup(title).split()


def words_to_title(words):
    return normalize_text(" ".join(words))


def find_column(df, *candidates):
    column_map = {str(col).strip().lower(): col for col in df.columns}
    for candidate in candidates:
        column = column_map.get(candidate.strip().lower())
        if column is not None:
            return column
    return None


def validate_template_columns(df_template):
    missing_columns = [
        column for column in REQUIRED_TEMPLATE_COLUMNS
        if column not in df_template.columns
    ]
    if missing_columns:
        raise ValueError(
            f"{TEMPLATE_FILE} 缺少必要列：{', '.join(missing_columns)}"
        )


def validate_required_columns(df, source_name):
    title_col = find_column(df, "Title", "title")
    handle_col = find_column(df, "Handle", "handle")
    source_id_col = find_column(df, "ID", "id")

    missing_columns = []
    if not title_col:
        missing_columns.append("Title")
    if not handle_col:
        missing_columns.append("Handle")

    if missing_columns:
        raise ValueError(
            f"{source_name} 缺少必要列：{', '.join(missing_columns)}"
        )

    return title_col, handle_col, source_id_col


def load_collection_frames(base_path):
    collection_dir = os.path.join(base_path, COLLECTION_DIR)
    single_file = os.path.join(base_path, DATA_FILE)

    csv_paths = []
    if os.path.isdir(collection_dir):
        csv_paths = sorted(
            os.path.join(collection_dir, name)
            for name in os.listdir(collection_dir)
            if name.lower().endswith(".csv")
        )

    if csv_paths:
        print(f"检测到 collection 文件夹，准备合并 {len(csv_paths)} 个 CSV 文件...")
        return csv_paths

    if os.path.exists(single_file):
        print(f"未检测到 collection 文件夹中的 CSV，改为使用单文件 {DATA_FILE} ...")
        return [single_file]

    raise FileNotFoundError(
        f"请确保目录内存在 {COLLECTION_DIR} 文件夹并包含 CSV，或存在 {DATA_FILE}"
    )


def read_and_merge_collection_data(base_path):
    csv_paths = load_collection_frames(base_path)
    normalized_frames = []

    for csv_path in csv_paths:
        try:
            df = read_csv_compat(csv_path)
        except Exception as exc:
            raise ValueError(f"读取文件失败：{csv_path}，原因：{exc}") from exc

        title_col, handle_col, source_id_col = validate_required_columns(
            df, os.path.basename(csv_path)
        )

        current = df.copy()
        current["Title"] = current[title_col]
        current["Handle"] = current[handle_col]
        if source_id_col:
            current["ID"] = current[source_id_col]
        elif "ID" not in current.columns:
            current["ID"] = ""
        current["_source_file"] = os.path.basename(csv_path)
        normalized_frames.append(current)
        print(f"已读取: {os.path.basename(csv_path)}，行数 {len(current)}")

    merged = pd.concat(normalized_frames, ignore_index=True, sort=False)

    # 随机拆分后常见的是同一行跨文件重复，这里仅按 Title + Handle 去重。
    before_dedup = len(merged)
    merged["Title"] = merged["Title"].apply(normalize_text)
    merged["Handle"] = merged["Handle"].fillna("").astype(str).str.strip()

    blank_title_count = int((merged["Title"] == "").sum())
    if blank_title_count:
        print(f"警告：已跳过 {blank_title_count} 行空 Title 目录数据")
        merged = merged[merged["Title"] != ""]

    blank_handle_count = int((merged["Handle"] == "").sum())
    if blank_handle_count:
        print(f"警告：已跳过 {blank_handle_count} 行空 Handle 目录数据，避免生成空集合链接")
        merged = merged[merged["Handle"] != ""]

    if merged.empty:
        raise ValueError("collection 数据中没有有效目录行，请检查 Title 和 Handle 列")

    merged = merged.drop_duplicates(subset=["Title", "Handle"], keep="first").reset_index(drop=True)
    deduped_count = before_dedup - len(merged)

    print(f"合并完成，总行数 {before_dedup}，去重后 {len(merged)}")
    if deduped_count:
        print(f"检测并移除了 {deduped_count} 条重复目录行（按 Title + Handle）")

    return merged, "ID"


def find_smart_parent(current_title, title_map):
    parts = title_words(current_title)
    if len(parts) <= 1:
        return 0

    for i in range(len(parts) - 1, 0, -1):
        potential_parent = words_to_title(parts[:i])
        if potential_parent not in title_map:
            continue

        next_word = parts[i].lower()
        if next_word in CONJUNCTIONS:
            continue

        return title_map[potential_parent]

    return 0


def build_title_map(titles, ids):
    title_map = {}
    duplicate_titles = set()

    for title, menu_id in zip(titles, ids):
        normalized_title = normalize_title_for_lookup(title)
        if not normalized_title:
            continue
        if normalized_title in title_map:
            duplicate_titles.add(title)
            continue
        title_map[normalized_title] = menu_id

    return title_map, duplicate_titles


def is_valid_parent_boundary(words, prefix_length):
    if prefix_length <= 0 or prefix_length >= len(words):
        return False

    last_word = words[prefix_length - 1].lower()
    next_word = words[prefix_length].lower()
    return last_word not in CONJUNCTIONS and next_word not in CONJUNCTIONS


def find_existing_parent_word_count(words, existing_titles):
    for i in range(len(words) - 1, 0, -1):
        potential_parent = words_to_title(words[:i])
        if potential_parent in existing_titles:
            return i
    return 0


def infer_missing_group_titles(titles):
    clean_titles = [normalize_text(title) for title in titles if normalize_text(title)]
    existing_titles = {
        normalize_title_for_lookup(title)
        for title in clean_titles
        if normalize_title_for_lookup(title)
    }

    prefix_children = defaultdict(set)
    title_word_lists = []

    for title in clean_titles:
        words = title_words(title)
        if len(words) <= 2:
            continue

        title_key = words_to_title(words)
        title_word_lists.append(words)

        for prefix_length in range(1, len(words)):
            if not is_valid_parent_boundary(words, prefix_length):
                continue

            prefix_title = words_to_title(words[:prefix_length])
            if prefix_title in existing_titles:
                continue

            parent_word_count = find_existing_parent_word_count(
                words[:prefix_length],
                existing_titles,
            )
            # 只补已有父级下面缺失的中间层级，避免把 "Skin Care Products"
            # 误拆出一个不存在的顶级集合 "Skin Care"。
            if parent_word_count == 0:
                continue
            if prefix_length - parent_word_count < MIN_INFERRED_GROUP_WORDS:
                continue

            prefix_children[tuple(words[:prefix_length])].add(title_key)

    candidate_prefixes = {
        prefix_words: len(children)
        for prefix_words, children in prefix_children.items()
        if len(children) >= MIN_INFERRED_GROUP_CHILDREN
    }

    inferred_groups = set()
    for words in title_word_lists:
        matching_prefixes = [
            prefix_words
            for prefix_words in candidate_prefixes
            if len(prefix_words) < len(words)
            and tuple(words[:len(prefix_words)]) == prefix_words
        ]
        if matching_prefixes:
            best_prefix = max(
                matching_prefixes,
                key=lambda prefix_words: (
                    candidate_prefixes[prefix_words],
                    len(prefix_words),
                ),
            )
            inferred_groups.add(words_to_title(best_prefix))

    return sorted(inferred_groups, key=str.casefold)


def build_title_handle_map(df_data):
    title_handle_map = {}
    for title, handle in zip(df_data["Title"], df_data["Handle"]):
        title_key = normalize_title_for_lookup(normalize_text(title))
        handle_text = normalize_text(handle)
        if title_key and title_key not in title_handle_map:
            title_handle_map[title_key] = handle_text

    return title_handle_map


def find_parent_handle_for_inferred_title(title, title_handle_map):
    words = title_words(title)
    for i in range(len(words) - 1, 0, -1):
        parent_title = words_to_title(words[:i])
        parent_handle = title_handle_map.get(parent_title, "")
        if parent_handle:
            return parent_handle

    return ""


def append_inferred_group_rows(df_data):
    inferred_titles = infer_missing_group_titles(df_data["Title"].tolist())
    if not inferred_titles:
        return df_data, [], []

    title_handle_map = build_title_handle_map(df_data)
    inferred_rows = []
    skipped_titles = []
    for title in inferred_titles:
        inherited_handle = find_parent_handle_for_inferred_title(title, title_handle_map)
        if not inherited_handle:
            skipped_titles.append(title)
            continue

        row = {column: "" for column in df_data.columns}
        row["Title"] = title
        row["Handle"] = inherited_handle
        if "_source_file" in row:
            row["_source_file"] = "auto-inferred-menu-group"
        inferred_rows.append(row)

    if not inferred_rows:
        return df_data, [], skipped_titles

    inferred_df = pd.DataFrame(inferred_rows, columns=df_data.columns)
    combined = pd.concat([df_data, inferred_df], ignore_index=True, sort=False)
    inserted_titles = [row["Title"] for row in inferred_rows]
    return combined, inserted_titles, skipped_titles


def trim_parent_prefix(current_title, parent_title):
    current = normalize_text(current_title)
    parent = normalize_text(parent_title)
    if not parent:
        return current
    prefix = f"{parent} "
    if current.startswith(prefix):
        return current[len(prefix):].strip()
    return current


def clamp_depth(result, max_depth=MAX_DEPTH):
    parent_map = dict(
        zip(
            result["Menu Item: ID"].astype(int),
            result["Menu Item: Parent ID"].fillna(0).astype(int),
        )
    )

    def get_level(menu_id, memo, visiting):
        if menu_id in memo:
            return memo[menu_id]
        if menu_id in visiting:
            memo[menu_id] = 1
            return 1

        visiting.add(menu_id)
        parent_id = parent_map.get(menu_id, 0)
        if parent_id == 0 or parent_id not in parent_map:
            level = 1
        else:
            level = 1 + get_level(parent_id, memo, visiting)
        visiting.remove(menu_id)
        memo[menu_id] = level
        return level

    while True:
        memo = {}
        changed = False

        for menu_id in list(parent_map):
            if get_level(menu_id, memo, set()) <= max_depth:
                continue
            parent_id = parent_map.get(menu_id, 0)
            parent_map[menu_id] = parent_map.get(parent_id, 0)
            changed = True

        if not changed:
            break

    result["Menu Item: Parent ID"] = (
        result["Menu Item: ID"].map(parent_map).fillna(0).astype(int)
    )


def validate_output_rows(result):
    id_values = result["Menu Item: ID"].fillna("").astype(str).str.strip()
    duplicate_ids = id_values[id_values.duplicated()].unique().tolist()
    if duplicate_ids:
        raise ValueError(f"生成结果中存在重复 Menu Item: ID：{', '.join(duplicate_ids)}")

    parent_ids = set(result["Menu Item: ID"].fillna(0).astype(int).tolist())
    invalid_parent_rows = result[
        ~result["Menu Item: Parent ID"].fillna(0).astype(int).isin(parent_ids | {0})
    ]
    if not invalid_parent_rows.empty:
        preview = ", ".join(
            invalid_parent_rows["Menu Item: Title"].fillna("").astype(str).head(5)
        )
        raise ValueError(f"生成结果中存在无效父级 ID，示例：{preview}")

    blank_handle_rows = result[
        result["Menu Item: Resource Handle"].fillna("").astype(str).str.strip() == ""
    ]
    if not blank_handle_rows.empty:
        preview = ", ".join(
            blank_handle_rows["Menu Item: Title"].fillna("").astype(str).head(10)
        )
        raise ValueError(
            "生成结果中存在空 Menu Item: Resource Handle，"
            f"请检查 collection Handle。示例：{preview}"
        )


def process_menu():
    base_path = get_application_path()
    input_template_file = os.path.join(base_path, TEMPLATE_FILE)
    output_file = os.path.join(base_path, OUTPUT_FILE)

    print("正在启动程序...")

    if not os.path.exists(input_template_file):
        print(f"错误：请确保文件夹内包含 {TEMPLATE_FILE}")
        return

    try:
        df_data, source_id_col = read_and_merge_collection_data(base_path)
        df_template = read_csv_compat(input_template_file)
        validate_template_columns(df_template)
    except Exception as exc:
        print(f"读取文件失败: {exc}")
        return

    df_data, inferred_titles, skipped_inferred_titles = append_inferred_group_rows(df_data)
    if inferred_titles:
        print(
            f"已自动补齐 {len(inferred_titles)} 个缺失的中间目录："
            + "，".join(inferred_titles)
        )
    if skipped_inferred_titles:
        print(
            f"警告：有 {len(skipped_inferred_titles)} 个推断目录因找不到可继承 Handle 已跳过："
            + "，".join(skipped_inferred_titles)
        )

    print("正在对合并后的目录数据按标题进行字母升序排列...")
    df_data = df_data.copy()
    df_data["_sort_title"] = df_data["Title"].fillna("").astype(str).str.casefold()
    df_data = (
        df_data.sort_values(by="_sort_title", ascending=True, kind="stable")
        .drop(columns="_sort_title")
        .reset_index(drop=True)
    )

    print("正在初始化菜单结构...")
    result = pd.DataFrame(index=df_data.index, columns=df_template.columns)

    if not df_template.empty:
        template_row = df_template.iloc[0]
        for col in result.columns:
            result[col] = template_row[col]
    else:
        template_row = pd.Series(dtype=object)

    full_titles = df_data["Title"].apply(normalize_text)
    handles = df_data["Handle"].fillna("").astype(str).str.strip()

    result["Menu Item: Title"] = full_titles
    result["Menu Item: Resource Handle"] = handles

    if source_id_col and "ID" in df_data.columns:
        result["ID"] = df_data["ID"]

    seq_ids = list(range(1, len(result) + 1))
    result["Row #"] = seq_ids
    result["Menu Item: ID"] = seq_ids

    print("正在计算层级关系...")
    title_map, duplicate_titles = build_title_map(
        full_titles.tolist(),
        result["Menu Item: ID"].tolist(),
    )
    result["Menu Item: Parent ID"] = full_titles.apply(
        lambda title: find_smart_parent(title, title_map)
    ).astype(int)

    if duplicate_titles:
        duplicate_text = ", ".join(sorted(normalize_text(title) for title in duplicate_titles))
        print(f"警告：检测到重复标题，父级识别按首次出现处理：{duplicate_text}")

    print(f"正在检查并限制层级深度（最大 {MAX_DEPTH} 级）...")
    clamp_depth(result, max_depth=MAX_DEPTH)

    print("正在优化标题显示...")
    full_title_map = dict(zip(result["Menu Item: ID"], full_titles))
    result["Menu Item: Title"] = [
        trim_parent_prefix(title, full_title_map.get(parent_id, ""))
        for title, parent_id in zip(full_titles, result["Menu Item: Parent ID"])
    ]

    display_title_map = dict(zip(result["Menu Item: ID"], result["Menu Item: Title"]))
    if "Menu Item: Parent Title" in result.columns:
        result["Menu Item: Parent Title"] = (
            result["Menu Item: Parent ID"].map(display_title_map).fillna("")
        )

    result["Menu Item: URL"] = handles.apply(
        lambda handle: f"/collections/{handle}" if handle else ""
    )

    try:
        validate_output_rows(result)
    except Exception as exc:
        print(f"输出校验失败: {exc}")
        return

    exclude = {
        "ID",
        "id",
        "Row #",
        "Menu Item: ID",
        "Menu Item: Parent ID",
        "Menu Item: Title",
        "Menu Item: Resource Handle",
        "Menu Item: Parent Title",
        "Menu Item: URL",
    }
    for col in result.columns:
        if col in exclude or col not in template_row.index:
            continue
        if result[col].isna().all():
            result[col] = template_row[col]

    try:
        result.to_csv(output_file, index=False, encoding="utf-8-sig")
        print("=" * 40)
        print("处理成功！")
        print("1. 已自动读取 collection 文件夹下全部 CSV 并合并")
        print("2. 合并后按 Title 字母升序统一排序")
        print("3. 自动去除跨文件重复目录行（按 Title + Handle）")
        print("4. 统一重建父子层级并限制目录深度不超过 3 层")
        print(f"结果已保存至: {output_file}")
        print("=" * 40)
    except Exception as exc:
        print(f"保存失败: {exc}")


if __name__ == "__main__":
    process_menu()
    try:
        input("按回车键退出程序...")
    except EOFError:
        pass
