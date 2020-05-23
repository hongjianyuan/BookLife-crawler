import requests  # 使用requsts套件
from bs4 import BeautifulSoup
import time
import re
import pandas as pd

all = []
page = 2
book_name_list = []
author = []
publisher = []
data = []
date = []
price = []
links = []
headers = {
    'user-agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36'}
while page < 5:
    url = "https://www.books.com.tw/web/books_nbtopm_01/?o=1&v=1&page=" + \
        str(page)
    re = requests.get(url, headers=headers)  # 將網頁資料GET下來
    re.encoding = 'utf8'
    soup = BeautifulSoup(re.text, "html.parser")  # 將網頁資料以html.parser
    # 获取书名
    book_names = soup.select(".msg h4")
    for book_name in book_names:
        book_name_list.append(book_name.text)
    # 获取作者名字与出版社
    elems = soup.select(".info a")
    for elem in elems:
        all.append(elem.text)
    for i in range(int(len(all))):
        if i % 2 == 0:
            author.append(all[i])
        else:
            publisher.append(all[i])
    # 获取书本资料
    data_elems = soup.select(".txt_cont  p")
    for data_elem in data_elems:
        data.append(data_elem.text.rstrip("more"))
    #  #获取价格
    price_elems = soup.select(".set2 ")
    for price_elem in price_elems:
        price.append(price_elem.text)
    # 获取link
    link_elems = soup.select("h4 a")
    for link_elem in link_elems:
        links.append(link_elem['href'])
    # 获取出版社日期
    date_elems = soup.select(".info span")
    for date_elem in date_elems:
        a = date_elem.text
        b = a.split("：")
        date.append(b[1])

    time.sleep(1)
    page += 1


# Create a Pandas dataframe from the data.
df1 = pd.DataFrame(book_name_list)
df2 = pd.DataFrame(author)
df3 = pd.DataFrame(publisher)
df4 = pd.DataFrame(data)
df5 = pd.DataFrame(date)
df6 = pd.DataFrame(price)
df7 = pd.DataFrame(links)

df = pd.concat([df1, df2, df3, df4, df5, df6, df7], axis=1)
# df.to_excel('books.xls')
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer1 = pd.ExcelWriter('Books.xlsx', engine='xlsxwriter')  # 改檔名
# # Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer1, sheet_name='Sheet1')
# Close the Pandas Excel writer and output the Excel file.
writer1.save()
