import re
import time
from tqdm import tqdm

from bs4 import BeautifulSoup
from selenium import webdriver										# webdriver の情報
from selenium.webdriver.common.by import By							# html の タブの情報を取得
from selenium.webdriver.common.keys import Keys						# キーボードを叩いた時に web ブラウザに情報を送信する
from selenium.webdriver.support import expected_conditions as EC	# 次にクリックしたページがどんな状態になっているかチェックする
from selenium.webdriver.support.ui import WebDriverWait				# 待機時間を設定
from selenium.webdriver.common.alert import Alert					# 確認ダイアログ制御


import load


"""
Selenium IE_Driver 仕様の準備
https://qiita.com/tsuttie/items/372f5d4cad37650711f1

https://qiita.com/__init__/items/1a90bfbbd2b1015bb90d

"""

SYSTEM_PATAM_FILEPATH = 'system_param.json'

SEARCH_PARAM_FILEPATH = '../お店探索条件.json'
INPUT_EXCEL_FILEPATH = '../DrinkingPartySheet.xlsm'



class Assister:

	def __init__(self):
		
		self.search_param = load.LoadJsonFile(SEARCH_PARAM_FILEPATH)
		self.system_param = load.LoadJsonFile(SYSTEM_PATAM_FILEPATH)

		self.member_dict = load.LoadMemberStatus(INPUT_EXCEL_FILEPATH, self.search_param['PartyName'])

		self.driver = webdriver.Ie(self.system_param['3rdParty']['IeDriver'])


	def LoadWebPageTop(self):
		
		self.driver.get(self.system_param['HotPepper']['Url'])
		time.sleep(1)
	

	def SetShopConditions(self):
		
		area_index = self.system_param['HotPepper']['Area'][self.search_param['Area']]
		prefecture_index = self.system_param['HotPepper']['Prefecture'][self.search_param['Prefecture']]
		self.driver.find_element_by_xpath(self.system_param['HotPepper']['Json']['PrefectureButton'].format(area_index, prefecture_index)).click()
		time.sleep(1)

		self.driver.find_element_by_xpath(self.system_param['HotPepper']['Json']['CityInput']).send_keys(self.search_param['City'])
		self.driver.find_element_by_xpath(self.system_param['HotPepper']['Json']['CityButton']).click()
		time.sleep(1)
	

	# 店の名前とURLを取得
	def GetShopDetail(self):

		def GetShopDetailTop(find_soup):
			shop_details = []
			
			for soup in find_soup:
				shop_soup = soup.find('h3', {'class': 'shopDetailStoreName'}).find('a')
				name = re.sub(r'[\u3000]', ' ', shop_soup.get_text())

				url = '{}{}'.format(self.system_param['HotPepper']['BaseURL'], shop_soup.get('href'))

				shop_details.append({
					'Name': name,
					'URL': url
				})
			return shop_details


		shop_details = []
		page_index = 0

		while True:
			page_soup = BeautifulSoup(self.driver.page_source, 'html5lib')			
			shop_details.extend(GetShopDetailTop(page_soup.find_all('div', {'class': 'shopDetailTop PR shopDetailWithCourseCalendar'})))
			shop_details.extend(GetShopDetailTop(page_soup.find_all('div', {'class': 'shopDetailTop shopDetailWithCourseCalendar'})))

			if len(shop_details) >= self.search_param['SearchNum']:
				shop_details = shop_details[:self.search_param['SearchNum']]
				break

			return_button_index = 10 if page_index == 0 else 11
			self.driver.find_element_by_xpath(self.system_param['HotPepper']['Json']['NextButton'].format(len(shop_soup) + 2, return_button_index)).click()

			time.sleep(1)

			page_index = page_index + 1
		
		return shop_details

	
	# 店のコース情報を取得
	def GetCourseDetail(self, shop_details):
		
		course_details = []

		"""
		for shop_elm in tqdm(shop_soups):
			
			try:
				name_elm = shop_elm.find_all('h3', {'class', 'shopDetailStoreName'})
				name_soup = BeautifulSoup(str(name_elm), 'html5lib')
				name_str_elm = name_soup.find_all(href=re.compile('/str'))[0]

				link = 'https://www.hotpepper.jp{}'.format(name_str_elm.get('href'))

				budget_elm = shop_elm.find_all('li', {'class', 'shopDetailInfoBudget'})
				budget_soup = BeautifulSoup(str(budget_elm), 'html5lib')

				capacity_elm = shop_elm.find_all('li', {'class', 'shopDetailInfoCapacity'})
				capacity_soup = BeautifulSoup(str(capacity_elm), 'html5lib')

				time_elm = shop_elm.find_all('li', {'class', 'shopDetailInfoTime'})
				time_soup = BeautifulSoup(str(time_elm), 'html5lib')
				
				name = re.sub(r'[\u3000]', ' ', name_str_elm.text)
				budget = re.sub(r'[\u3000\n\t\[\]]', '', budget_soup.text)
				budget = re.sub(r'[\u3000]', ' ', budget)
				capacity = re.sub(r'[\u3000\n\t\[\]]', '', capacity_soup.text)
				capacity = re.sub(r'[\u3000]', ' ', capacity)
				time = re.sub(r'[\n\t\[\]]', '', time_soup.text)
				time = re.sub(r'[\u3000]', ' ', time)


				course_details = self.GetCourseDetail(link)


				shop_details.append({
					'Name'    : name,
					'Budget'  : budget,
					'Capacity': capacity,
					'Time'    : time,
					'Link'    : link,
					'Courses' : course_details
				})
			
			except:
				pass
		"""

		return course_details

	



if __name__ == '__main__':
	
	assist = Assister()
	
	assist.LoadWebPageTop()
	assist.SetShopConditions()

	shop_details = assist.GetShopDetail()
	course_details = assist.GetCourseDetail(shop_details)

	
#	for detail in shop_details:
#		print('-'*30)
#		print(detail)
#	print('-'*30)
