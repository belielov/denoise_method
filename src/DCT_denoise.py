import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib import rcParams
from scipy.fftpack import dct, idct

# 输入文件路径和输出文件路径
input_path = r'10-15M-2.xlsx'
output_path = r'10-15M-DCT_denoised.xlsx'

# 读取Excel文件
data = pd.read_excel(
    input_path,   # Excel 文件路径
    header=None,  # 不将任何行作为列名
    names=['Wave_number', 'Intensity']  # 指定列名
)

# 提取横、纵坐标
x = data['Wave_number'].values
y = data['Intensity'].values

# ----- DCT稀疏分解去噪逻辑 ------
def dct_denoise(data, threshold_radio=0.1, mode='hard'):
    """
    DCT稀疏分解去噪
    步骤：
        1. 对信号进行DCT变换，得到频域系数
        2. 通过阈值保留主要系数，去除噪声相关的小系数
        3. 逆DCT变换重构信号
    Args：
        data: 输入的待平滑数据
        threshold_radio: 百分比阈值
        mode: 阈值处理模式
    Returns:
        y_smooth: 经过DCT稀疏分解去噪后的 Intensity 值
    """
    # DCT变换
    dct_coeffs = dct(data, norm='ortho')  # 使用正交归一化，确保变换的正交性，避免逆变换时需额外缩放。

    # 计算阈值
    sorted_coeffs = np.sort(np.abs(dct_coeffs))[::-1]  # 降序排序DCT系数绝对值
    idx = int(len(sorted_coeffs) * threshold_radio)  # 保留前 threshold_radio% 的系数
    threshold = sorted_coeffs[idx] if idx < len(sorted_coeffs) else 0

    # 阈值处理
    if mode == 'hard':
        dct_coeffs_thresh = dct_coeffs * (np.abs(dct_coeffs) >= threshold)  # 保留绝对值 ≥ 阈值的系数，其余置零。
    elif mode == 'soft':
        dct_coeffs_thresh = np.sign(dct_coeffs) * np.maximum(np.abs(dct_coeffs) - threshold, 0)  # 软阈值需保留系数符号，确保信号相位不变。
    else:
        raise ValueError("模式需为 'hard' 或 'soft'")

    # 逆DCT重构信号
    y_smooth = idct(dct_coeffs_thresh, norm='ortho')
    return y_smooth

# 应用DCT变换（保留前10%的系数，使用硬阈值）
y_smooth = dct_denoise(y, threshold_radio=0.07, mode='soft')

# 将处理后的数据存储为DataFrame
smoothed_data = pd.DataFrame({'Wave_number': x, 'Intensity': y_smooth})

# 保存数据到Excel
smoothed_data.to_excel(output_path, index=False, header=True)

# ----- 绘图与字体设置 -----
rcParams['font.sans-serif'] = ['SimHei']
rcParams['axes.unicode_minus'] = False

plt.figure(figsize=(10, 6))
plt.plot(x, y, label='原始数据', color='blue', alpha=0.6)
plt.plot(x, y_smooth, label='DCT去噪数据', color='red', linewidth=2)
plt.xlabel('Wave_number')
plt.ylabel('Intensity')
plt.title('原始数据与DCT去噪后的数据')
plt.legend(loc='best')
plt.grid(True)
plt.show()

print(f'处理后的数据已保存到：{output_path}')