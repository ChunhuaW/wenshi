import time
def main():
    import pandas as pd
    from pyecharts import options as opts
    from pyecharts.charts import Tree
    import random

    pattern_df = pd.read_excel('纹饰.xlsx')
    links_df = pd.read_excel('links纹饰&瓷器（1）.xlsx')
    # 从纹饰表创建name到type的映射字典
    type_dict = pattern_df.set_index('name')['type'].to_dict()
    # 使用 map 方法将 links 表中的 target 替换为 type
    links_df['value'] = links_df['target'].map(type_dict)

    # 统计每种纹饰种类下的陶瓷和纹饰数量
    df_group = links_df.groupby('value').agg({
        'source': 'count',
        'target': 'nunique'
    }).rename(columns={'source': '陶瓷数量', 'target': '纹饰数量'}).reset_index()

    # 将数据转换为树形结构的数据格式
    tree_data = [
        {
            "name": "纹饰类型",
            "children": []
        }
    ]

    for index, row in df_group.iterrows():
        value_node = {
            "name": str(row['value']),
            "children": []
        }
        # 当点击对应陶瓷数量时，随机展示三个不同的陶瓷
        if row['陶瓷数量'] > 0:
            all_sources = links_df[links_df['value'] == row['value']]['source'].unique()
            random_sources = random.sample(list(all_sources), min(3, len(all_sources)))
            count_node = {
                "name": f'陶瓷数量: {row["陶瓷数量"]}',
                "children": []
            }
            for source in random_sources:
                count_node["children"].append({"name": source})
            value_node["children"].append(count_node)
        # 当点击对应纹饰数量时，随机展示三个纹饰
        if row['纹饰数量'] > 0:
            all_targets = links_df[links_df['value'] == row['value']]['target'].unique()
            random_targets = random.sample(list(all_targets), min(3, len(all_targets)))
            count_node = {
                "name": f'纹饰数量: {row["纹饰数量"]}',
                "children": []
            }
            for target in random_targets:
                count_node["children"].append({"name": target})
                # 根据纹饰数量设置线条宽度
                count_node['lineStyle'] = {'width': row['纹饰数量'] * 0.3}
                count_node['lineStyle']['color'] = '#EBCE98'
            value_node["children"].append(count_node)
        tree_data[0]["children"].append(value_node)

    # 创建树形图对象
    tree = (
        Tree(init_opts=opts.InitOpts(width="800px", height="1000px"))
            .add("", tree_data)
            .set_global_opts(
            title_opts=opts.TitleOpts(title="纹饰可视化"),
        )
            .set_series_opts(
            label_opts=opts.LabelOpts(font_family="黑体", font_size=14)
        )
    )
    # 生成本地网页链接
    tree.render("index.html")
if __name__ == "__main__":
    while True:
        main()
        time.sleep(1)
