import os
import xlrd
import requests

def download_files_from_xls_by_colnames(
    xls_path, 
    save_dir, 
    filename_col_names,  # 集合，包含所有需要下载的url列名
    start_row=0, 
    end_row=None
):
    """
    根据xls文件的指定行数范围，下载指定字段（列名）提供的图片、mp3等文件

    :param xls_path: xls文件路径
    :param save_dir: 文件保存目录
    :param filename_col_names: 需要下载的url列名集合（表头集合）
    :param start_row: 起始行（包含，默认0，包含表头）
    :param end_row: 结束行（包含，若为None则到最后一行）
    """
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)
    workbook = xlrd.open_workbook(xls_path)
    sheet = workbook.sheet_by_index(0)
    num_rows = sheet.nrows
    header = sheet.row_values(0)
    # 获取所有列名对应的索引
    colname_to_idx = {}
    for colname in filename_col_names:
        try:
            colname_to_idx[colname] = header.index(colname)
        except ValueError:
            print(f"未找到列名: {colname}")
    if not colname_to_idx:
        print("未找到任何有效的列名，退出。")
        return

    if end_row is None or end_row >= num_rows:
        end_row = num_rows - 1

    for row_idx in range(start_row, end_row + 1):
        row = sheet.row_values(row_idx)
        for colname, col_idx in colname_to_idx.items():
            url = str(row[col_idx]).strip()
            if not url or url.lower() == "none":
                print(f"第{row_idx+1}行，列[{colname}]无有效url，跳过")
                continue
            # 文件名从url自动获取
            filename = os.path.basename(url.split("?")[0])
            save_path = os.path.join(save_dir, filename)
            try:
                resp = requests.get(url, timeout=15)
                if resp.status_code == 200:
                    with open(save_path, "wb") as f:
                        f.write(resp.content)
                    print(f"已下载: {filename}")
                else:
                    print(f"下载失败: {filename}，状态码: {resp.status_code}")
            except Exception as e:
                print(f"下载{url}出错: {e}")


import xlrd

# 示例用法
if __name__ == "__main__":
    # 假设xls文件路径为"diandu/diandu.xlsx"
    # 假设需要下载的url列名集合为{"originImgUrl", "originSoundUrl"}
    xls_path = "diandu/SecondGrade/diandu.xlsx"
    save_dir = "diandu/SecondGrade/downloads"
    filename_col_names = {"originImgUrl", "originSoundUrl"}  # 请根据实际表头修改
    start_row = 1  # 包含表头
    end_row = -1
    if end_row == -1 :
        workbook = xlrd.open_workbook(xls_path)
        sheet = workbook.sheet_by_index(0)
        end_row = sheet.nrows  # 自动读取xlsx总行数，end_row为最后一行索引
    download_files_from_xls_by_colnames(
        xls_path, 
        save_dir, 
        filename_col_names, 
        start_row, 
        end_row
    )
