from urllib.parse import quote
from urllib.request import urlopen
from bs4 import BeautifulSoup
import time
import logging
from collections import Iterable
import os
import random
import xlwt
import xlrd
from xlutils.copy import copy
import re

logging.basicConfig(level=logging.INFO)



# 创建Excel

def CreateExcel(path, sheets, title):
  try:
      logging.info('创建Excel: %s' % path)
      book = xlwt.Workbook()
      for sheet_name in sheets:
          sheet = book.add_sheet(sheet_name, cell_overwrite_ok=True)
          for index, item in enumerate(title):
              sheet.write(0, index, item, set_style('Times New Roman', 220, True))
      book.save(path)
  except IOError:
      return '创建Excel出错！'



# 设置Excel样式

def set_style(name, height, bold=False):
  style = xlwt.XFStyle()  # 初始化样式

  font = xlwt.Font()  # 为样式创建字体
  font.name = name  # 'Times New Roman'
  font.bold = bold
  font.color_index = 4
  font.height = height

  # borders= xlwt.Borders()
  # borders.left= 6
  # borders.right= 6
  # borders.top= 6
  # borders.bottom= 6

  style.font = font
  # style.borders = borders

  return style



# 加载Excel获得副本

def LoadExcel(path):
  logging.info('加载Excel：%s' % path)
  book = xlrd.open_workbook(path)
  copy_book = copy(book)
  return copy_book



# 判断内容是否存在

def ExistContent(book, sheet_name):
  sheet = book.get_sheet(sheet_name)
  if len(sheet.get_rows()) >= 2:
      return True
  else:
      return False



# 写入Excel并保存]\

def WriteToTxcel(book, sheet_name, content, path):
  logging.info('%s 数据写入到 (%s-%s)' % (sheet_name, os.path.basename(path), sheet_name))
  sheet = book.get_sheet(sheet_name)
  for index, item in enumerate(content):
      for sub_index, sub_item in enumerate(item):
          sheet.write(sub_index + 1, index, sub_item)
  book.save(path)



# 获得页面景点信息

def GetPageSite(url):
  try:
      page = urlopen(url)
  except AttributeError:
      logging.info('抓取失败！')
      return 'ERROR'
  try:
      bs_obj = BeautifulSoup(page.read(), 'lxml')
      # 不存在页面
      if len(bs_obj.find('div', {'class': 'result_list'}).contents) <= 0:
          logging.info('当前页面没有信息！')
          return 'NoPage'
      else:
          page_site_info = bs_obj.find('div', {'class': 'result_list'}).children
  except AttributeError:
      logging.info('访问被禁止！')
      return None
  return page_site_info





#获取页面数目
def GetPageNumber(url):
    try:
        page = urlopen(url)
    except AttributeError:
        logging.info('抓取失败')
        return 'ERROR'
    try:
        bs_obj = BeautifulSoup(page.read(), 'lxml')
        if len(bs_obj.find('div',{'class':'result_list'}).contents) <= 0:
            logging.info('当前页面无信息')
            return 'NoPage'
        else:
            page_site_info = bs_obj.find('div', {'class':'pager'}).get_text()
    except AttributeError:
        logging.info('访问禁止')
        return None

     #提取页面数
    page_num = re.findall(r'\d+\.?\d*', page_site_info.split('...')[-1])

    return int(page_num[0])

# 去除重复数据

def FilterData(data):
  return list(set(data))



# 格式化获取信息

def GetItem(site_info):
  site_items = {}  # 储存景点信息
  site_info1 = site_info.attrs
  site_items['name'] = site_info1['data-sight-name']  # 名称
  site_items['position'] = site_info1['data-point']  # 经纬度
  site_items['address'] = site_info1['data-districts'] + ' ' + site_info1['data-address']  # 地理位置
  site_items['sale number'] = site_info1['data-sale-count']  # 销售量

  site_level = site_info.find('span', {'class': 'level'})
  if site_level:
      site_level = site_level.get_text()
  site_hot = site_info.find('span', {'class': 'product_star_level'})
  if site_hot:
      site_hot = site_info.find('span', {'class': 'product_star_level'}).em.get_text()
      site_hot = site_hot.split(' ')[1]

  site_price = site_info.find('span', {'class': 'sight_item_price'})
  if site_price:
      site_price = site_info.find('span', {'class': 'sight_item_price'}).em.get_text()

  site_items['level'] = site_level
  site_items['site_hot'] = site_hot
  site_items['site_price'] = site_price

  return site_items



# 获取一个省的所有景点

def GetProvinceSite(province_name):
  site_name = quote(province_name)  # 处理汉字问题
  url1 = 'http://piao.qunar.com/ticket/list.htm?keyword='
  url2 = '&region=&from=mps_search_suggest&page='
  url = url1 + site_name + url2

  NAME = []  # 景点名称
  POSITION = []  # 坐标
  ADDRESS = []  # 地址
  SALE_NUM = []  # 票销量
  SALE_PRI = []  # 售价
  STAR = []  # 景点星级
  SITE_LEVEL = []  # 景点热度

  i = 0  # 页面
  page_num = GetPageNumber(url + str(i + 1))  # 页面数
  logging.info('当前城市 %s 存在 %s 个页面' % (province_name, page_num))
  flag = True  # 访问非正常退出标志
  while i < page_num:  # 遍历页面
      i = i + 1
      # 随机暂停1--5秒，防止访问过频繁被服务器禁止访问
      time.sleep(1 + 4 * random.random())

      # 获取网页信息
      url_full = url + str(i)
      print(url_full)
      site_info = GetPageSite(url_full)
      # 当访问被禁止的时候等待一段时间再进行访问
      while site_info is None:
          wait_time = 60 + 540 * random.random()
          while wait_time >= 0:
              time.sleep(1)
              logging.info('访问被禁止，等待 %s 秒钟后继续访问' % wait_time)
              wait_time = wait_time - 1
          # 继续访问
          site_info = GetPageSite(url_full)
      if site_info == 'NoPage':  # 访问完成
          logging.info('当前城市 %s 访问完成，退出访问！' % province_name)
          break
      elif site_info == 'ERROR':  # 访问出错
          logging.info('当前城市 %s 访问出错,退出访问' % province_name)
          flag = False
          break
      else:
          # 返回对象是否正常
          if not isinstance(site_info, Iterable):
              logging.info('当前页面对象不可迭代 ，跳过 %s' % i)
              continue
          else:
              # 循环获取页面信息
              for site in site_info:
                  info = GetItem(site)
                  NAME.append(info['name'])
                  POSITION.append(info['position'])
                  ADDRESS.append(info['address'])
                  SALE_NUM.append(info['sale number'])
                  SITE_LEVEL.append(info['site_hot'])
                  SALE_PRI.append(info['site_price'])
                  STAR.append(info['level'])

              logging.info('当前访问城市 %s,取到第 %s 组数据: %s' % (province_name, i, info['name']))

  return flag, NAME, POSITION, ADDRESS, SALE_NUM, SALE_PRI, STAR, SITE_LEVEL


def ProvinceInfo(province_path):
  tlist = []
  with open(province_path, 'r', encoding='utf-8') as f:
      lines = f.readlines()
      for line in lines:
          tlist = line.split('，')
  return tlist



# 生成Json格式文本

def GenerateJson(ExcelPath, JsonPath, TransPos=False):
  try:
      if os.path.exists(JsonPath):
          os.remove(JsonPath)
      json_file = open(JsonPath, 'a', encoding='utf-8')
      book = xlrd.open_workbook(ExcelPath)
  except IOError as e:
      return e
  sheets = book.sheet_names()
  for sheet_name in sheets[0:1]:
      sheet = book.sheet_by_name(sheet_name)
      row_0 = sheet.row_values(0, 0, sheet.ncols - 1)  # 标题栏数据
      # 获得热度栏数据
      for indx, head in enumerate(row_0):
          if head == '销售量':
              index = indx
              break
      level = sheet.col_values(index, 1, sheet.nrows - 1)

      if not TransPos:
          for indx, head in enumerate(row_0):
              if head == '经纬度':
                  index = indx
                  break
          pos = sheet.col_values(index, 1, sheet.nrows - 1)
          for i, p in enumerate(pos):
              if int(level[i]) > 0:
                  lng = p.split(',')[0]
                  lat = p.split(',')[1]
                  lev = level[i]
                  json_temp = '{"lng":' + str(lng) + ',"lat":' + str(lat) + ', "count":' + str(lev) + '}, '
                  json_file.write(json_temp + '\n')
      else:
          pass
  json_file.close()
  return 'TransPos=%s,Trans pos to json done.' % TransPos


if __name__ == '__main__':
  excel_path = r'g:\python program examples\Python-master/Info.xls'
  province_path = r'g:\python program examples\Python-master/info.txt'
  # Excel表头信息
  title = ['名称', '经纬度', '地址', '销售量', '起售价', '星级', '热度']

  # 加载省份列表
  province_list = ProvinceInfo(province_path)
  # 如果旧表不存在则创建新表
  if not os.path.exists(excel_path):
      CreateExcel(excel_path, province_list, title)

  # 爬取内容
  book = LoadExcel(excel_path)
  for index, province in enumerate(province_list):
      # 判断内容是否存在,存在则跳过当前城市
      if ExistContent(book, province):
          logging.info('当前访问城市 %s,该内容存在，跳过' % province)
          continue
      # 获取城市的景点信息
      Contents = GetProvinceSite(province)
      if Contents[0]:  # 获取正常则保存
          WriteToTxcel(book, province, Contents[1:], excel_path)