import re
import urllib.request
import xlwt
from bs4 import BeautifulSoup
import time

def askURL(url):
    # 设置‘请求头’
    # 以谷歌浏览器为例，输入在网址栏“chrome://version/”可以进入浏览器内核
    # 从而查看用户代理信息，即‘请求头’中的“User-agent”信息内容
    header = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)"
                      "Chrome/131.0.0.0 Safari/537.36"
    }
    # 构造请求头的对象
    req = urllib.request.Request(url=url, headers=header)
    html = ""
    # 发送请求获取响应的内容
    try:
        response = urllib.request.urlopen(req)
        html = response.read().decode("utf-8")
        # 防止访问请求过于频繁，触发反爬虫机制
        time.sleep(2)
    # 捕获异常信息
    except urllib.error.URLError as e:
        if hasattr(e, 'code'):
            print(e.code)
        if hasattr(e, 'reason'):
            print(e.reason)
    return html

# 影片详细链接的正则表达式
findlink = re.compile(r'<a href="(.*?)">')
# 影片图片的正则表达式
findImgSrc = re.compile(r'<img.*src="(.*?)"', re.S)
# 影片片名的正则表达式
findTitle = re.compile(r'span class="title">(.*)</span>')
# 影片评分的正则表达式
findRating = re.compile(r'<span class="rating_num" property="v:average">(.*)</span>')  # 修正此处
# 评论人数的正则表达式
findJudge = re.compile(r'<span>(\d*)人评价</span>')
# 概况的正则表达式
findInq = re.compile(r'<span class="inq">(.*)</span>')
# 影片相关内容
findBD = re.compile(r'<p class="">(.*?)</p>', re.S)

def getData(baseurl):
    datalist = []
    for i in range(0, 10):
        url = baseurl + "?start=" + str(i * 25)
        html = askURL(url)
        if not html:
            continue  # 如果没有获取到网页内容，则跳过该页
        soup = BeautifulSoup(html, "html.parser")
        # 查找符合的内容并且形成一个列表
        for item in soup.find_all('div', class_="item"):
            data = []
            item = str(item)
            # 影片详细的链接
            link = re.findall(findlink, item)[0]
            data.append(link)
            # 图片详细的链接
            imgSrc = re.findall(findImgSrc, item)[0]
            data.append(imgSrc)
            # 影片片名
            titles = re.findall(findTitle, item)
            # 判断是否有英文原文
            if len(titles) == 2:
                ctitle = titles[0]
                data.append(ctitle)
                # 去掉特殊字符
                otitle = titles[1].replace("/", "")
                data.append(otitle)
            else:
                data.append(titles[0])
                # 只有一个中文译名时，需要将英文原名的位置用空格占位
                data.append(" ")
            # 影片评分
            rating = re.findall(findRating, item)[0]
            data.append(rating)
            # 评论人数
            judgeNum = re.findall(findJudge, item)[0]
            data.append(judgeNum)
            # 影片概况
            inq = re.findall(findInq, item)
            if len(inq) != 0:
                # 去掉句号
                inq = inq[0].replace("。", "")
                data.append(inq)
            else:
                # 没有影片概况则用空格占位，保证格式的一致性
                data.append(" ")
            # 影片相关内容
            bd = re.findall(findBD, item)
            # 去掉换行符，引号外写上r识别‘\s’
            bd = re.sub(r'<br(\s+)?/>(\s+)?', " ", bd[0] if bd else "")
            bd = re.sub('/', " ", bd)  # 去掉'/'
            data.append(bd.strip())  # 去掉前后空格
            datalist.append(data)
    return datalist  # 确保返回数据

def saveData(datalist, savepath):
    print('正在保存数据')
    # 创建Workbook对象，并且设置为utf-8格式
    book = xlwt.Workbook(encoding="utf-8", style_compression=0)
    # 创建一个sheet工作表
    sheet = book.add_sheet('豆瓣电影Top250', cell_overwrite_ok=True)
    # 设置每列的列名
    col = ("电影详细链接", "图片链接", "影片中文名", "影片外国名", "评分", "评价数", "概况", "相关信息")
    # 第一行书写列名
    for i in range(0, 8):
        sheet.write(0, i, col[i])
    # 写入相关数据到表格中（此处输入的数据量仅设为30）
    for i in range(0, 250):
        print(f"第{i + 1}条")
        data = datalist[i]
        # 嵌入内循环
        for j in range(0, 8):
            sheet.write(i + 1, j, data[j])
    # 保存文件到指定路径
    book.save(savepath)

if __name__ == "__main__":
    # 豆瓣电影top250排行榜url地址
    baseurl = "https://movie.douban.com/top250"
    # 抓取网页
    datalist = getData(baseurl)
    # 保存数据的路径
    savepath = "./豆瓣电影Top 250.xls"
    # 保存数据
    saveData(datalist, savepath)

