import os
import pandas as pd

def txt_to_xlsx(input_folder, output_folder):
    """
    将指定文件夹中的所有 .txt 文件转换为 .xlsx 文件，跳过以 # 开头的行。

    参数:
    - input_folder: 存放 .txt 文件的文件夹路径。
    - output_folder: 存放 .xlsx 文件的目标文件夹路径。
    """
    # 如果目标文件夹不存在，创建它
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 遍历输入文件夹中的所有文件
    for file_name in os.listdir(input_folder):
        if file_name.endswith('.txt'):  # 检查是否是 .txt 文件
            input_path = os.path.join(input_folder, file_name)
            output_path = os.path.join(output_folder, file_name.replace('.txt', '.xlsx'))

            # 读取 .txt 文件
            try:
                # 尝试用空格或逗号作为分隔符读取文件，并跳过以 # 开头的行
                df = pd.read_csv(input_path, delimiter=',', header=None, comment='#',encoding='gbk')
                # 将数据写入 Excel
                df.to_excel(output_path, index=False, header=False)
                print(f"成功转换: {file_name} -> {os.path.basename(output_path)}")
            except Exception as e:
                print(f"转换失败: {file_name}，错误: {e}")

# 设置输入和输出文件夹路径
input_folder = "D:\\桌面\\TXT"  # 替换为存放 .txt 文件的路径
output_folder = "D:\\桌面\\Excel"  # 替换为存放 .xlsx 文件的目标路径

# 调用函数
txt_to_xlsx(input_folder, output_folder)
