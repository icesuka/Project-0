import requests
import re
from bs4 import BeautifulSoup
import xlwt

def main():
    baseurl = 'http://www.woshipm.com/archive/page/'
    datalist = getData(baseurl)
    savepath = "wspm.xls"
    saveData(datalist,savepath)

leixing = re.compile(r'rel="category tag">(.*?)</a>',re.S)
biaoti = re.compile(r'<a aria-label="(.*?)"',re.S)
wangzhi = re.compile(r'<a aria-label=".*?" title=".*?" href="(.*?)"',re.S)
riqi = re.compile(r'<time itemprop="datePublished">(.*?)</time>',re.S)


def getData(baseurl):
    datalist = []
    for i in range(1,26):
        url = baseurl + str(i)
        html = askURL(url)

        soup = BeautifulSoup(html, "html.parser")
        for item in soup.find_all('article', itemscope='itemscope'):
            data = []
            item = str(item)

            lx = re.findall(leixing, item)
            data.append(lx)

            bt = re.findall(biaoti, item)
            data.append(bt)

            wz = re.findall(wangzhi, item)
            data.append(wz)

            rq = re.findall(riqi, item)
            data.append(rq)
            datalist.append(data)

    return datalist


def askURL(url):
    headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.83 Safari/537.36'}
    resp = requests.get(url, headers=headers)
    resp.encoding = 'utf-8'
    html=resp.text
    return html

def saveData(datalist,savepath):
    print("save....")
    book = xlwt.Workbook(encoding="utf-8",style_compression=0)
    sheet = book.add_sheet('wspm',cell_overwrite_ok=True)
    col = ("类型","标题","网址","日期")
    for i in range(0,4):
        sheet.write(0,i,col[i])
    for i in range(0,500):
        print("第%d条" %(i+1))
        data = datalist[i]
        for j in range(0,4):
            sheet.write(i+1,j,data[j])

    book.save(savepath)

if __name__ == "__main__":
    main()
    print("爬取完毕！")