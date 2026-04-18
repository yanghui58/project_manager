import pandas as pd
from tqdm import tqdm


# 定义人物关系统计函数
def count_relationships(name_show, persons_relation):
    """
    统计人物关系出现的次数，确保只出现一次（如AB，不出现BA）
    """
    for i in range(len(name_show)):
        for j in range(i + 1, len(name_show)):
            # 确保人物对的顺序一致（按字母顺序排序）
            person1, person2 = sorted([name_show[i], name_show[j]])
            name1name2 = f'{person1},{person2}'
            if name1name2 not in persons_relation:
                # 如果关系不在字典中，则初始化为 0
                persons_relation[name1name2] = 0
            # 关系出现次数加 1
            persons_relation[name1name2] += 1
    return persons_relation


def get_name_aliases():
    name_aliases = {}
    with open("孔乙己人物名单.txt", encoding='utf-8') as f:
        d = f.readlines()
        for d_in in d:
            name = d_in.strip().split(" ")[0]
            name_aliases[name] = []
    with open("孔乙己人物别称映射字典.txt", encoding='utf-8') as f:
        d = f.readlines()
        for d_in in d:
            d_in = d_in.strip().split(" ")
            for d_1 in d_in:
                if d_1 in name_aliases:
                    break
            for d_2 in d_in:
                if d_2 != d_1:
                    name_aliases[d_1].append(d_2)
    for name in name_aliases:
        if len(name_aliases[name]) == 0:
            name_aliases[name].append(name)
    return name_aliases


def get_ta(content, prev_content, ta, name_aliases, persons_count, name_show):
    """
    统计单数代词（他/她）指代的人物
    """
    ind = 0
    while True:
        try:
            ind = content.index(ta, ind)
            ind += 1
        except:
            break
        content_ = content[:ind] if not prev_content else prev_content + "\n" + content[:ind]
        content_ = content_[::-1]
        max_ind, max_name = 0, ""
        for name in name_aliases:
            for name_in in name_aliases[name]:
                name_in = name_in[::-1]
                if name_in in content_:
                    if content_.index(name_in) > max_ind:
                        max_ind = content_.index(name_in)
                        max_name = name
        if max_name != "":
            if max_name not in persons_count:
                persons_count[max_name] = 0
            persons_count[max_name] += 1
            if max_name not in name_show:
                name_show.append(max_name)


def get_tamen(content, prev_content, tamen, name_aliases, persons_count, name_show):
    """
    统计复数代词（他们/她们）指代的人物
    """
    ind = 0
    while True:
        try:
            ind = content.index(tamen, ind)
            ind += 1
        except:
            break
        content_ = content[:ind] if not prev_content else prev_content + "\n" + content[:ind]
        for name in name_aliases:
            for name_in in name_aliases[name]:
                if name_in in content_:
                    if name not in persons_count:
                        persons_count[name] = 0
                    persons_count[name] += 1
                    if name not in name_show:
                        name_show.append(name)


if __name__ == '__main__':
    # 读取人物别名映射
    name_aliases = get_name_aliases()
    persons_count = {}
    persons_relation = {}

    # 指定单个txt文件路径
    file_path = r'D:\py\pythonProject\pythonProject\《孔乙己》.txt'

    # 读取文件内容
    with open(file_path, 'r', encoding='utf-8') as f:
        contents = f.readlines()

    print(f"正在分析文件: {file_path}")
    print(f"文件总行数: {len(contents)}")

    # 处理文件内容
    for i in tqdm(range(len(contents)), desc="处理进度"):
        name_show = []
        content = contents[i]
        prev_content = contents[i - 1] if i > 0 else ""

        # 处理单数代词
        get_ta(content, prev_content, '他', name_aliases, persons_count, name_show)
        get_ta(content, prev_content, '她', name_aliases, persons_count, name_show)

        # 处理复数代词
        get_tamen(content, prev_content, '他们', name_aliases, persons_count, name_show)
        get_tamen(content, prev_content, '她们', name_aliases, persons_count, name_show)

        # 统计直接出现的人名
        for name in name_aliases:
            if name not in persons_count:
                persons_count[name] = 0
            for name_in in name_aliases[name]:
                cnt = content.count(name_in)
                persons_count[name] += cnt
                if cnt > 0 and name not in name_show:
                    name_show.append(name)

        # 统计当前行的人物关系
        persons_relation = count_relationships(name_show, persons_relation)

    # 保存结果到CSV文件
    persons_count_df = pd.DataFrame([[x, persons_count[x]] for x in persons_count], columns=['姓名', '出场次数'])

    # 处理关系数据，确保人物1和人物2的顺序与保存的一致
    relations_data = []
    for relation_key, weight in persons_relation.items():
        person1, person2 = relation_key.split(",")
        relations_data.append([person1, person2, weight])

    persons_relation_df = pd.DataFrame(relations_data, columns=['人物1', '人物2', 'weight'])

    persons_count_df.to_csv("孔乙己-出场次数.csv", index=None, encoding='utf-8-sig')
    persons_relation_df.to_csv("孔乙己-人物关系.csv", index=None, encoding='utf-8-sig')

    print("\n分析完成!")
    print(f"人物出场次数已保存到:孔乙己-出场次数.csv")
    print(f"人物关系网络已保存到:孔乙己-人物关系.csv")

    # 显示统计摘要
    print(f"\n统计摘要:")
    print(f"共识别到 {len(persons_count)} 个不同人物")
    print(f"共识别到 {len(persons_relation)} 组人物关系")

    # 按出场次数排序并显示前3名
    top_characters = persons_count_df.sort_values('出场次数', ascending=False).head(3)
    print("\n出场次数最多的前3位人物:")
    for idx, row in top_characters.iterrows():
        print(f"  {row['姓名']}: {row['出场次数']}次")

    # 显示最重要的3组关系
    print("\n最重要的3组人物关系:")
    persons_relation_df_sorted = persons_relation_df.sort_values('weight', ascending=False).head(3)
    for idx, row in persons_relation_df_sorted.iterrows():
        print(f"  {row['人物1']} - {row['人物2']}: {row['weight']}次")