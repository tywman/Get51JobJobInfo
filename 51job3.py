# -*- coding:utf-8 -*-
import urllib.request
import urllib
import re
from bs4 import BeautifulSoup
import xlwt#用来创建excel文档并写入数据
import time

def getCityCode(cityname):
    citydic = {'深圳':'040000',
               '北京':'010000',
               '上海':'020000'}
    if citydic[cityname] == None:
        citycode = '000000'
    else:
        citycode = citydic[cityname]
    return citycode

#获取原码
def get_content(page,keywords,city):
    s = urllib.parse.quote(keywords)
    url ='http://search.51job.com/list/%s,000000,0000,00,9,99,%s,2,%s.html'%(getCityCode(city),s,str(page))
    a = urllib.request.urlopen(url)#打开网址
    html = a.read().decode('gbk')#读取源代码并转为unicode
#    soup = BeautifulSoup(html,'xml')
    return html

def get_pages(html):
    pages = re.compile(r'<span class="td">共(\d*?)页')
    totalpages = re.findall(pages,html)
    return int(totalpages[0])
    
def get(html):
#    resultlists = []
#    results = soup.find_all('div', {'class': 'dw_table'})
#    for result in results:
#        items = result.find_all('div',{'class': 'el'})
#        for item in items:
#            titles = item.find_all('p',{'class':'t1 '})
#            companys = item.find_all('span',{'class':'t2'})
#            locations = item.find_all('span',{'class':'t3'})
#            moneys = item.find_all('span',{'class':'t4'})
#            publics = item.find_all('span',{'class':'t5'})
#            for title, company,location,money,public in zip(titles, companys,locations,moneys,publics):
#                print('%s,%s,%s,%s,%s,%s'%(title.find('a')['title'], company.get_text(),location.get_text(),money.get_text(),public.get_text(),title.find('a')['href']))
#                resultlist = [title.find('a')['title'], company.get_text(),location.get_text(),money.get_text(),public.get_text(),title.find('a')['href']]
#                resultlists.append(resultlist)
    reg = re.compile(r'class="t1 ">.*? <a target="_blank" title="(.*?)".*? href="(.*?)".*?<span class="t2"><a target="_blank" title="(.*?)".*?<span class="t3">(.*?)</span>.*?<span class="t4">(.*?)</span>.*? <span class="t5">(.*?)</span>',re.S)#匹配换行符
    resultlists=re.findall(reg,html)
    return resultlists

def split_(value):
    if value.find('-'):
        valuelow = float(value[:value.find('-')])
        valuehig = float(value[value.find('-')+1:])
    else:
        valuelow = float(value);
        valuehig = float(value);
    return valuelow,valuehig

def excel_write(items,index):
    #爬取到的内容写入excel表格
    for item in items:#职位信息
        zw = item[0]
        gs = item[2]
        qy = item[3]
        cs = qy.find('-')
        if cs > 0:
            cs = qy[:cs]
        else:
            cs = qy
        xz = item[4]
        if xz.find('千/月') > 0:
            valuelow,valuehig = split_(xz[:xz.find('千/月')])
            valuelow = valuelow*12/10
            valuehig = valuehig*12/10
        elif(xz.find('万/月') > 0):
            valuelow,valuehig = split_(xz[:xz.find('万/月')])
            valuelow = valuelow*12
            valuehig = valuehig*12
        elif(xz.find('万/年') > 0):
            valuelow,valuehig = split_(xz[:xz.find('万/年')])
        elif(xz.find('100万以上') > 0):
            valuelow = 100
            valuehig = 100
        else:
            valuelow = 0
            valuehig = 0
        
        rq = item[5]
        lj = item[1]
        ws.write(index,0,xlwt.Formula('HYPERLINK("%s";"%s")'%(lj,zw)))
        ws.write(index,1,gs)
        ws.write(index,2,cs)
        ws.write(index,3,qy)
        ws.write(index,4,xz)
        ws.write(index,5,rq)
        ws.write(index,6,valuelow)
        ws.write(index,7,valuehig)
#        for i in range(len(item)):
#            #print item[i]
#            ws.write(index,i,item[i])#行，列，数据
##        print(index)
        index+=1

keyword = '物联网'
city = '深圳'
newTable=keyword + ".xls"#表格名称
wb = xlwt.Workbook(encoding='utf-8')#创建excel文件，声明编码
ws = wb.add_sheet('sheet1')#创建表格
headData = ['招聘职位','公司','地址','区域','薪资','日期']#表头部信息
for colnum in range(len(headData)):
    ws.write(0, colnum, headData[colnum], xlwt.easyxf('font: bold on'))  # 行，列

pages = get_pages(get_content(1,keyword,city))
start_row = 1
for each in range(pages):
    print('total %s页,当期处理第%d页'%(pages,each))
    results = get(get_content(each,keyword,city))
    excel_write(results,start_row)
    print('共%d条数据，起始行：%d'%(len(results),start_row))
    start_row += len(results)
    time.sleep(1)

wb.save(newTable)