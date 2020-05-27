# 导入通用模块
import pandas as pd
import re
import os


def kml(rootdir):
    """
    批量提取文件夹中的kml文件中的地理位置信息
    ----
    Arguments:
    ----
        rootdir {str} -- 指定目录，即包含kml文件的文件夹的路径
    """

    lists = os.listdir(rootdir)

    # 初始化正则表达式模式
    longitude = u"(经度：.+?<{1})"  # 查找的经度：和后续数字
    latitude = u"(纬度：.+?<{1})"  # 查找的纬度
    time = u"(时间：.+?<{1})"  # 查找时间
    altitude = u"(海拔：.+?<{1})"  # 查找海拔

    # 遍历目录文件夹下的所有kml文件，并提取地理位置信息
    infolists = []  # 初始化一个列表，储存包含kml信息的列表
    for i in range(0, len(lists)):
        path = os.path.join(rootdir, lists[i])
        output = list()  # 初始化一个列表，储存每个kml中信息
        if os.path.isfile(path):
            f = open(path, "r")
            txt = f.read()
            header_name = os.path.basename(path)
            output.append(header_name)  # 添加文件名
            pattern = re.compile(longitude)  # 匹配经度
            results = pattern.findall(txt)  # 匹配经度
            for result in results:
                output.append(result)
            pattern = re.compile(latitude)
            results = pattern.findall(txt)
            for result in results:
                output.append(result)
            pattern = re.compile(time)
            results = pattern.findall(txt)
            for result in results:
                output.append(result)
            pattern = re.compile(altitude)
            results = pattern.findall(txt)
            for result in results:
                output.append(result)
            f.close()
        infolists.append(output)
    df = pd.DataFrame(infolists)
    return df


if __name__ == "__main__":
    df = kml("包含kml文件的文件夹")
    print(df)
