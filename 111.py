from excelChange import txt_to_excel_with_history

import os


def get_all_files(folder_path):
    """
    读取指定文件夹内的所有文件名（包括子文件夹中的文件）

    参数:
        folder_path: 文件夹路径

    返回:
        包含所有文件路径的列表
    """
    # 检查文件夹是否存在
    if not os.path.exists(folder_path):
        print(f"错误：文件夹 '{folder_path}' 不存在")
        return []

    if not os.path.isdir(folder_path):
        print(f"错误：'{folder_path}' 不是一个文件夹")
        return []

    file_list = []
    # 遍历文件夹及其子文件夹
    for root, dirs, files in os.walk(folder_path):
        # 输出当前文件夹
        print(f"\n文件夹: {root}")
        print("-" * (len(root) + 4))

        # 输出文件夹中的文件
        for file in files:
            file_path = os.path.join(root, file)
            file_list.append(file_path)
            print(f"文件: {file}")

        # 输出子文件夹
        for dir in dirs:
            print(f"子文件夹: {dir}")

    return file_list


# 在这里输入你要读取的文件夹路径
# folder = input("请输入要读取的文件夹路径: ")
# 或者直接指定路径，例如：
folder = "D:\\1112"

files = get_all_files(folder)
print(files)


for i in files:
    txt_to_excel_with_history(i, "角色数据_历史追踪.xlsx", group="A组")