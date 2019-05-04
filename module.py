#!/usr/bin/env python
# coding: utf-8

# In[4]:


import requests
from bs4 import BeautifulSoup
import openpyxl

book_div = "mod type02_m012 clearfix"

def fetch_books_info(url, kind, dump_files=False, show_info=True):
    html = requests.get(url).text
    soup = BeautifulSoup(html, "html.parser")
    try:
        pages = int(soup.select(".cnt_page span")[0].text)
        print(" Total: ", pages, " pages")
        dumps = []
        for page in range(1, pages+1):
            pageurl = url + "&page=" + str(page).strip()
            print(" The", page, " page", pageurl)
            dump = fetch_page(pageurl, kind, dump_files, show_info)
            dumps += dump
        if dump_files is True:
            return dumps
    except:
        dump = fetch_page(url, kind, dump_files, show_info)
        if dump_files is True:
            return dump

def fetch_page(url, kind, dump_files=False, show_info=True):
    html = requests.get(url).text
    soup = BeautifulSoup(html, 'html.parser')
    res = soup.find_all("div", {"class":book_div})[0]
    items = res.select(".item")
    dump = []
    for item in items:
        msg = item.select(".msg")[0]
        src = item.select("a img")[0]["src"]
        title = msg.select("a")[0].text
        imgurl = src.split("?i=")[-1].split("&")[0]
        author = msg.select("a")[1].text
        publish = msg.select("a")[2].text
        date = msg.find("span").text.split("：")[-1]
        onsale = item.select(".price .set2")[0].text
        content = item.select(".txt_cont")[0].text.replace(" ","").strip()
        if show_info is True:
            print("\n分類:" + kind)
            print("書名:" + title)
            print("圖片網址:" + imgurl)
            print("作者:" + author)
            print("出版社:" + publish)
            print("出版日期:" + date)
            print("優惠價:" + onsale)
            print("內容:" + content)
            list_data = [kind, title, imgurl, author, publish, date, onsale, content]
            dump.append(list_data)
    
    if dump_files is True:
        return dump
    return

def save_dumps(filename, dump):
    workbook = openpyxl.Workbook()
    sheet = workbook.worksheets[0]
    
    list_title=["分類","書名","圖片網址","作者","出版社","出版日期","優惠價","內容"]
    sheet.append(list_title)
    for item in dump:
        sheet.append(item)
    
    workbook.save(filename + ".xlsx")


# In[ ]:




