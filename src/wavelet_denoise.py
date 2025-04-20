import numpy as np
import pandas as pd
import pywt
from matplotlib import rcParams, pyplot as plt

# 输入文件路径和输出文件路径
input_path = r'10-15M-2.xlsx'
output_path = r'10-15M-wavelet_denoised.xlsx'

# 读取 Excel 文件
data = pd.read_excel(
    input_path,  # Excel 文件路径
    skiprows=1,  # 跳过文件中的第一行
    header= None,  # 不将任何行作为列名
    names=['Wave_number', 'Intensity']  # 指定列名
)

# 提取横、纵坐标
x = data['Wave_number'].values
y = data['Intensity'].values

# ---- 小波软阈值去噪逻辑 ----
def wavelet_denoise(data, wavelet='db4', level=3, mode='soft'):
    """ 小波分解
    小波变换将信号分解为不同频率的成分。通过多级分解，信号被逐层拆分为：
        - 近似系数（Approximation Coefficients）：低频部分，反映信号的主要趋势或轮廓。
        - 细节系数（Detail Coefficients）：高频部分，包含细节信息（如噪声或突变）。
    coeffs 的结构:
        - 调用 pywt.wavedec() 后，coeffs 是一个列表，其结构如下：
        - coeffs = [approx_levelN, detail_levelN, detail_levelN-1, ..., detail_level1]
        - approx_levelN：最高层级（level 指定）的近似系数（低频）。
        - detail_levelN 到 detail_level1：从最高层级到第1层的细节系数（高频）。
    """
    coeffs = pywt.wavedec(
        data,  # 输入数据
        wavelet,   # 小波基类型，影响分解的灵敏度和去噪效果
        mode='per',  # 信号扩展模式，'per'（周期扩展）避免边界效应
        level=level,  # 分解层级，决定分解的深度。层级越高，越关注低频部分。
    )

    """ 计算阈值（使用通用阈值）
    基于 噪声标准差估计 和 通用阈值（Universal Threshold） 方法
    
    目标1：估计噪声的标准差（σ）。
    关键步骤：
        1. 选择高频细节系数：
            - coeffs[-level] 表示小波分解中最底层的高频细节系数（如 level=3 时对应 detail3）。
            - 高频系数通常以噪声为主，适合用于估计噪声强度。
        2. 计算中值绝对偏差（MAD）：
            - np.median(np.abs(coeffs[-level])) 计算高频系数绝对值的中值。
            - 中值（Median）对异常值鲁棒，避免信号中的有效成分干扰噪声估计。
        3. 转换为标准差：
            - MAD（中值绝对偏差）与标准差的关系为：σ = MAD / 0.6745
            - 这里 0.6745 是标准正态分布下 MAD 与标准差的换算系数。
        4. 特殊情况处理：
            - if level > 0 else 1.0 表示当 level=0（未分解）时，默认噪声标准差为 1.0。
            
    目标2：根据噪声标准差和信号长度，计算去噪的软阈值。
    公式来源：
        - 通用阈值（Universal Threshold）由 Donoho 和 Johnstone 提出，公式为：
            T = σ * ⋅sqrt(2ln(N))
        其中 N 是信号长度，σ 是噪声标准差。
    数学意义：
        - 对高斯白噪声，此阈值能以高概率保证所有噪声系数被抑制。
        - sqrt(2ln(N))是理论推导的缩放因子，随信号长度增加缓慢增长。
    """
    sigma = (np.median(np.abs(coeffs[-level])) / 0.6745) if level > 0 else 1.0
    threshold = sigma * np.sqrt(2 * np.log(len(data)))

    # 软阈值处理
    coeffs_thresh = []
    for i, coeff in enumerate(coeffs):
        if i > 0:  # 仅处理细节系数
            coeffs_thresh.append(pywt.threshold(coeff, threshold, mode=mode))
        else:
            coeffs_thresh.append(coeff)

    # 小波重构
    y_smooth = pywt.waverec(coeffs_thresh, wavelet, mode='per')
    return y_smooth[:len(data)]  # 确保长度一致

# 应用小波去噪
y_smooth = wavelet_denoise(y, wavelet='db4', level=3, mode='soft')

# 将处理后的数据存储为DataFrame
smoothed_data = pd.DataFrame({'Wave_number': x, 'Intensity_smooth': y_smooth})

# 保存数据到 Excel
smoothed_data.to_excel(
    output_path,   # 参数1：输出文件的路径
    index=False,   # 参数2：不保存行索引
    header=True    # 参数3：保存列名作为标题行（第一行）
)

# ---- 绘图与字体设置 ----
rcParams['font.sans-serif'] = ['SimHei']
rcParams['axes.unicode_minus'] = False

plt.figure(figsize=(10, 6))
plt.plot(x, y, label='原始数据', color='blue', alpha=0.6)
plt.plot(x, y_smooth, label='小波去噪数据', color='red', linewidth=2)
plt.xlabel('Wave_number')
plt.ylabel('Intensity')
plt.title('原始数据与小波去噪后的数据')
plt.legend(loc='best')
plt.grid(True)
plt.show()

print(f'处理后的数据已保存到：{output_path}')
