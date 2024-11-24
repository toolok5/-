import os
import pandas as pd

# 定义文件夹路径
folder_path = r'C:\excel'  # 使用原始字符串（r''）避免路径中的转义字符

def process_csv(file_path):
    # 读取整个CSV文件
    df = pd.read_csv(file_path,encoding='gbk')

    # 提取1到52列的数据用于第一个工作表
    df_original = df.iloc[:, :52].copy()

    # 提取1到5列和53到62列的数据用于第二个工作表
    columns_to_keep = list(range(5)) + list(range(52, 62))
    df_processed = df.iloc[:, columns_to_keep].copy()

    # 定义输出文件名（保持原文件名，但扩展名为xlsx）
    output_file_path = os.path.splitext(file_path)[0] + '.xlsx'

    # 使用 ExcelWriter 写入数据
    with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='w') as writer:
        # 将原始数据写入第一个工作表
        df_original.to_excel(writer, sheet_name='KPI,MR处理', index=False)
        # 将处理后的数据写入第二个工作表
        df_processed.to_excel(writer, sheet_name='数据处理', index=False)

    return output_file_path


def process_files_in_folder():
    """遍历文件夹中的所有 CSV 文件并处理"""
    for filename in os.listdir(folder_path):
        if filename.endswith('.csv'):
            file_path = os.path.join(folder_path, filename)
            output_file_path = process_csv(file_path)
            print(f'Processed and saved {filename} as {os.path.basename(output_file_path)}')


def main():
    """
    统一入口，调用 process_files_in_folder()，用于外部调用
    """
    process_files_in_folder()


if __name__ == "__main__":
    main()
