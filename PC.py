# -*- codeing==UTF-8 -*-
# -*- author：刘新宇
# @Time:2020/5/28/23:55
# @File:PC.py
import re
import urllib.request
from bs4 import BeautifulSoup
import xlwt


def main():
    baseurl = "https://movie.douban.com/top250?start="
    date = getUrl(baseurl)
    print(date)
    savepath = "豆瓣250.xls"
    save(date, savepath)


def getHtml(url):
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                             "Chrome/83.0.4103.61 Safari/537.36"}
    request = urllib.request.Request(url, headers=headers)
    try:
        response = urllib.request.urlopen(request)
        return response.read().decode('utf-8')
    except urllib.error.URLError as e:
        if hasattr(e, "code"):
            print(e.code)
        if hasattr(e, "reason"):
            print(e.reason)


findlink = re.compile(r'href="(.*?)">')
findname = re.compile(r'<span class="title">(.*?)</span>')
findsorce = re.compile(r'<span class="rating_num" property="v:average">(.*?)</span>')


def getUrl(baseurl):
    datelist = []

    for i in range(0, 10):
        url = baseurl + str(i*25)
        html = getHtml(url)
        soup = BeautifulSoup(html, "html.parser")
        for item in soup.find_all("div", class_="item"):
            date = []
            item = str(item)
            link = re.findall(findlink, item)[0]
            date.append(link)
            name = re.findall(findname, item)[0]
            date.append(name)
            sorce = re.findall(findsorce, item)[0]
            date.append(sorce)
            datelist.append(date)
    return datelist


def save(datelist, SavePath):
    workbook = xlwt.Workbook(encoding='utf-8')  # 创建workbook对象
    worksheet = workbook.add_sheet("sheet1")  # 创建工作表
    col = ("电影路径", "电影名", "评分")
    for i in range(0, 3):
        worksheet.write(0, i, col[i])
    for i in range(0, 250):
        date = datelist[i]
        for j in range(0, 3):
            worksheet.write(i + 1, j, date[j])
        workbook.save(SavePath)


if __name__ == '__main__':
    main()
    print("保存完毕")
