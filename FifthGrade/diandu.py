import os
import json
import pandas as pd

def extract_piece_rows_from_json(json_file):
    """
    从单个json文件中提取pieces数组中的每个对象，合并page级别字段
    """
    rows = []
    try:
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)['Data']
        # 兼容单对象或数组
        if isinstance(data, dict):
            pages = [data]
        elif isinstance(data, list):
            pages = data
        else:
            # 尝试修正JSON格式并重新解析
            try:
                with open(json_file, 'r', encoding='utf-8') as f:
                    raw = json.load(f)['Data']
                # 可根据需要添加更多修正规则
                data = json.loads(raw)
                if isinstance(data, dict):
                    pages = [data]
                elif isinstance(data, list):
                    pages = data
                else:
                    print(f"{json_file} 文件格式不正确，无法解析为dict或list")
                    return rows
            except Exception as e:
                print(f"{json_file} 文件格式异常，修正后仍无法解析：{e}")
                return rows
        for page in pages:
            page_fields = {k: v for k, v in page.items() if k != "pieces"}
            pieces = page.get("pieces", [])
            for piece in pieces:
                row = {}
                # 展开page字段
                row.update(page_fields)
                # 展开piece字段
                for k, v in piece.items():
                    if isinstance(v, dict):
                        # coordinate等嵌套对象，展开为piece_coordinate_x等
                        for subk, subv in v.items():
                            row[f"{k}_{subk}"] = subv
                    else:
                        row[k] = v
                rows.append(row)
    except Exception as e:
        print(f"处理{json_file}时出错：", e)
    return rows

def diandu_jsons_to_xls(json_dir, xls_file):
    """
    循环读取指定目录下以diandu开头的json文件，提取pieces数组内容写入xls
    """
    all_rows = []
    for fname in os.listdir(json_dir):
        if fname.startswith("diandu") and fname.endswith(".json"):
            fpath = os.path.join(json_dir, fname)
            rows = extract_piece_rows_from_json(fpath)
            all_rows.extend(rows)
    if not all_rows:
        print("未找到有效数据")
        return
    # 统一所有字段
    all_columns = set()
    for row in all_rows:
        all_columns.update(row.keys())
    all_columns = sorted(all_columns)
    # 填充缺失字段
    for row in all_rows:
        for col in all_columns:
            if col not in row:
                row[col] = None
    df = pd.DataFrame(all_rows, columns=all_columns)
    try:
        df.to_excel(xls_file, index=False)
        print(f"已写入{xls_file}")
    except Exception as e:
        print("写入Excel时出错：", e)

if __name__ == "__main__":
    # 现在假设json文件在diandu目录下
    json_dir = "diandu\FifthGrade"
    xls_file = "diandu\FifthGrade\diandu.xlsx"
    diandu_jsons_to_xls(json_dir, xls_file)