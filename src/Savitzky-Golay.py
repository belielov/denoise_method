import pandas as pd
from scipy.signal import savgol_filter
import os
import matplotlib.pyplot as plt
from matplotlib import rcParams

# 输入文件路径和输出文件路径
input_path = r'10-15M-2.xlsx'  # 替换为实际文件路径
output_path = r'10-15M-2-Savitzky-Golay.xlsx'  # 输出路径

# 读取 Excel 文件，跳过第一行
data = pd.read_excel(input_path, skiprows=1, header=None, names=['x', 'y'])

# 提取横坐标和纵坐标
x = data['x'].values
y = data['y'].values

# 使用 Savitzky-Golay 滤波器平滑数据
# window_length: 滤波窗口大小（必须为奇数且小于数据点数量）
# polyorder: 多项式阶数
y_smooth = savgol_filter(y, window_length=17, polyorder=3)

# 将平滑后的数据存储为 DataFrame
smoothed_data = pd.DataFrame({'Wave_number': x, 'Intensity_smooth': y_smooth})

# 保存平滑后的数据到 Excel
smoothed_data.to_excel(output_path, index=False, header=True)

# 设置字体以支持中文
rcParams['font.sans-serif'] = ['SimHei']  # SimHei 是常见的黑体字体
rcParams['axes.unicode_minus'] = False  # 处理负号显示问题

# 绘制原始数据和平滑后的数据
plt.figure(figsize=(10, 6))
plt.plot(x, y, label='原始数据', color='blue', alpha=0.6)  # 原始数据
plt.plot(x, y_smooth, label='平滑数据', color='red', linewidth=2)  # 平滑后的数据
plt.xlabel('Wave_number')
plt.ylabel('Intensity')
plt.title('原始数据与Savitzky-Golay滤波后的数据')
plt.legend(loc='best')
plt.grid(True)

# 显示图像
plt.show()

print(f"平滑后的数据已保存到: {output_path}")
