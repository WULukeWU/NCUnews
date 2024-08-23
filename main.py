from selenium import webdriver
from chromedriver_py import binary_path
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import re
from bs4 import BeautifulSoup
import requests
import datetime
from google.oauth2.service_account import Credentials
import gspread
import pandas as pd
import os
import json

# ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓

# Variables - GitHub
line_notify_id = os.environ['LINE_NOTIFY_ID']
sheet_key = os.environ['GOOGLE_SHEETS_KEY']
gs_credentials = os.environ['GS_CREDENTIALS']
service = Service(ChromeDriverManager().install())


# Variables - Google Colab
# line_notify_id = LINE_NOTIFY_ID
# sheet_key = GOOGLE_SHEETS_KEY
# gs_credentials = GS_CREDENTIALS
# service = Service(binary_path)

# ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑

# LINE Notify ID
LINE_Notify_IDs = list(line_notify_id.split())

# 定義查找nid代碼函數
def find_nid(title, text):
    title_line_numbers = []
    for i, line in enumerate(text.split('\n')):
        if title in line:
            title_line_numbers.append(i)

    if not title_line_numbers:
        print(f'Cannot find "{title}" in the text.')
        return None

    title_line_number = title_line_numbers[0]
    title_line = text.split('\n')[title_line_number]

    nid_start_index = title_line.index('nid="') + 5
    nid_end_index = title_line.index('"', nid_start_index)
    nid = title_line[nid_start_index:nid_end_index]

    return nid

# 取得網頁內容
def get_content(url):
  # 發送GET請求獲取網頁內容
  response = requests.get(url)

  # 解析HTML內容
  soup = BeautifulSoup(response.content, 'html.parser')

  # 找到所有的 <p> 標籤
  p_tags = soup.find_all('p')

  # 整理文字內容
  text_list = []
  for p in p_tags:
      text = p.text.strip()
      text_list.append(text)
  text = ' '.join(text_list)
  text = ' '.join(text.split())  # 利用 split() 和 join() 將多個空白轉成單一空白
  # text = text.replace(' ', '\n')  # 將空白轉換成換行符號
  text = text.replace(' ', '')  # 刪除空白
  return text

text_limit = 1000-20

# LINE Notify
def LINE_Notify(category, date, department, title, link, content):

  send_info_1 = f'【{category}】{title}\n⦾公告單位：{department}\n⦾發佈日期：{date}'
  send_info_2 = f'⦾內容：' if content != '' else ''
  send_info_3 = f'⦾更多資訊：{link}'

  text_len = len(send_info_1) + len(send_info_2) + len(send_info_3)
  if content != '':
    if text_len + len(content) > text_limit:
      content = f'{content[:(text_limit - text_len)]}⋯'
    params_message = f'{send_info_1}\n{send_info_2}{content}\n{send_info_3}'
  else:
    params_message = f'{send_info_1}\n{send_info_3}'

  for LINE_Notify_ID in LINE_Notify_IDs:
    headers = {
            'Authorization': 'Bearer ' + LINE_Notify_ID,
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    params = {'message': params_message}

    r = requests.post('https://notify-api.line.me/api/notify',
                            headers=headers, params=params)
    print(r.status_code)  #200

# Google Sheets 紀錄
scope = ['https://www.googleapis.com/auth/spreadsheets']
info = json.loads(gs_credentials)

creds = Credentials.from_service_account_info(info, scopes=scope)
gs = gspread.authorize(creds)

def google_sheets_refresh():

  global sheet, worksheet, rows_sheets, df

  # 使用表格的key打開表格
  sheet = gs.open_by_key(sheet_key)
  worksheet = sheet.get_worksheet(0)

  # 讀取所有行
  rows_sheets = worksheet.get_all_values()
  # print(len(rows_sheets))
  # 使用pandas創建數據框
  df = pd.DataFrame(rows_sheets)

def main(url):

    # chromedriver 設定
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')

    driver = webdriver.Chrome(service=service, options=options)
    driver.get(url)

    # 等待網頁載入完成
    driver.implicitly_wait(10)

    # 打印整個頁面的 HTML 代碼
    page_source = driver.page_source
    # print(page_source)

    # 使用 BeautifulSoup 解析 HTML 內容
    soup = BeautifulSoup(page_source, 'html.parser')

    # 找到所有 class 為 'eventlist' 的 <li> 元素
    events = soup.find_all('li', class_='eventlist')

    # 初始化一個列表來存儲提取的數據
    event_data = []

    # 遍歷每個 'eventlist' <li> 元素
    for event in events:
        # 提取鏈接
        link = event.find('a')['href']
        # 提取活動日期
        date = event.find('div', class_='event-date').find('i').get_text(strip=True)
        # 提取活動標題
        title = event.find('div', class_='subject').find('h3').get_text(strip=True)
        # 提取活動類別
        category = event.find('u', class_='tabconstr').get_text(strip=True)
        # 提取部門名稱
        department = event.find('u', class_='tabDept').get_text(strip=True)
        
        # 將提取的數據存儲在一個字典中
        event_info = {
            'link': f'https://www.ncu.edu.tw/tw/events/{link}',
            'date': date,
            'title': title,
            'category': category,
            'department': department
        }
        
        # 將字典添加到列表中
        event_data.append(event_info)

    # 打印提取的活動數據
    # for event in event_data:
    #     print(event)

    # 定義需要查找的最新幾筆資料
    numbers_of_new_data = 10
    numbers_of_new_data = min(numbers_of_new_data, len(event_data))

    # 印出最新幾筆資料的標題、單位和連結
    for i in range(numbers_of_new_data - 1, -1, -1):

        event = event_data[i]

        # link_publish = f"{url[:url.find('ischool')]}ischool/public/news_view/show.php?nid={nid}"
        # link = f"{url[:url.find('ischool')]}ischool/public/news_view/show.php?nid={nid}"

        # 打開公告詳細頁面
        driver.get(event['link'])
        driver.implicitly_wait(10)

        # 抓取詳細頁面的 HTML 並打印
        detailed_page_html = driver.page_source
        # print("詳細頁面HTML:")
        # print(detailed_page_html)
        # print("---")

        # 使用 BeautifulSoup 解析 HTML
        soup = BeautifulSoup(detailed_page_html, 'html.parser')
        # print(soup)

        # Locate the div with the class "editor"
        editor_div = soup.find('div', class_='editor')

        # Extract the text
        if editor_div:
            content = editor_div.get_text(separator='\n', strip=True)
            event['content'] = content
            # print(content)
        else:
            # print("Content not found.")
            event['content'] = '-'

        print(f"title:{event['title']}\tcategory:{event['category']}\tdate:{event['date']}\tdepartment:{event['department']}\tlink:{event['link']}\tcontent:{event['content']}")

        # 獲取當前日期
        today = datetime.date.today()

        # 將日期格式化為2023/02/11的形式
        formatted_date = today.strftime("%Y/%m/%d")

        # 檢查nid是否已經存在於表格中
        sent = not(event['link'] in links)
        # print(sent, event['link'], links)

        if sent:

          # 檢查標題是否已經存在於表格中
          titles = df[4].tolist()
          if event['title'] in titles:
            continue

          # 獲取新行
          now = datetime.datetime.now() + datetime.timedelta(hours=8)
          new_row = [now.strftime("%Y-%m-%d %H:%M:%S"), event['category'], event['date'], event['department'], event['title'], event['link'], event['content']]

          # 將新行添加到工作表中
          worksheet.append_row(new_row)

          # 獲取新行的索引
          new_row_index = len(rows_sheets) + 1
          rows_sheets.append([])
          # print(new_row_index)

          # 更新單元格
          cell_list = worksheet.range('A{}:G{}'.format(new_row_index, new_row_index))
          for cell, value in zip(cell_list, new_row):
              cell.value = value
          worksheet.update_cells(cell_list)

          # 更新links列表
          links.append(event['link'])

          # 傳送至LINE Notify
          print(f"Sent: {event['link']}", end=' ')
          LINE_Notify(event['category'], event['date'], event['department'], event['title'], event['link'], event['content'])

        # 刪除nid
        # del link

    # 關閉網頁
    driver.quit()

# 開啟網頁
urls = [
    'https://www.ncu.edu.tw/tw/events/index.php',
]

if __name__ == "__main__":

  # 刷新Google Sheets表格
  google_sheets_refresh()

  # 取得Google Sheets nids列表
  _links = df[5].tolist()
  links = []
  for l in _links:
    try:
      links.append(l)
    except:
      continue

  error_links = []

  for url in urls:

    finished = False
    try_times_limit = 2
    for _ in range(try_times_limit):
      try:
        main(url)
        finished = True
        break
      except:
        print('retrying...')
        next

    if not finished:
      error_links.append(url)
      print(f'error : {url}')

  if len(error_links) == 0:
    print("--------------------------------------------------\nAll Finished Successfully. ")
  else:
    print(f"--------------------------------------------------\nAll Finished, Here Are All The Links That Cannot Be Sent Successfully. ({len(error_links)} files)")
    for error_link in error_links:
      print(error_link)