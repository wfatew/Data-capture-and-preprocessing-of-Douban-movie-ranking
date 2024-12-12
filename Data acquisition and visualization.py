import pandas as pd
import matplotlib.pyplot as plt

# 读取本地的数据
file = '豆瓣电影Top 250.xls'  # 确保文件路径正确
temp = pd.read_excel(file)
temp.head(10)
# 数据探索
temp.info()
# 统计缺失值
print("缺失值统计：")
print(temp.isnull().sum())
# 监测重复值
print("重复值统计：")
print(temp.duplicated().sum())
# 属性归纳
data = temp[['影片中文名', '评分', '评价数', '概况']]
# 根据评论数排序
data_sorted_by_comments = data.sort_values(by='评价数', ascending=False)
print("按评价数排序后的前10条数据：")
print(data_sorted_by_comments.head(10))
# 根据评分这一字段进行分组
group = data.groupby("评分")
for i in list(group):
    print(i)
# 统计每个评分组中的电影数量
group_count = group.count()
group_count.sort_index(inplace=True)
# 设置柱状图中显示的中文字体
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
# 绘制评价人数统计柱状图
x = data.iloc[::4, 0]  # 每四行数据取一次，避免数据过多导致图表拥挤
y = data.iloc[::4, 2]
plt.figure(figsize=(18, 6), dpi=250)
plt.bar(x, y, label='评价人数')
# 设置图像的各个元素
plt.xticks(rotation=60, rotation_mode='anchor', ha='right')  # 旋转X轴标签
plt.ylabel("评价人数/人", fontsize=15)
plt.xlabel("电影名称", fontsize=15)
plt.title("评价人数统计柱状图", fontsize=25)
plt.grid(linestyle=":")
plt.legend(loc="upper right")
# 显示图像
plt.show()

# 绘制各评分电影数量统计柱状图
plt.figure(figsize=(10, 4), dpi=250)
# 绘制每个评分下的电影数量
plt.bar(range(group_count.shape[0]), group_count['影片中文名'], label='电影数量', tick_label=group_count.index)
# 对图像进行设置
plt.ylabel("电影数量", fontsize=15)
plt.title("各分数电影数量统计柱状图", fontsize=20)
plt.grid(linestyle=":")
plt.legend(loc='upper right')
# 显示图像
plt.show()
