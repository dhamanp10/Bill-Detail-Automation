from django.shortcuts import render, redirect, HttpResponseRedirect
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from django.views import View
# Import mimetypes module
import mimetypes
# import os module
import os
# Import HttpResponse module
from django.http.response import HttpResponse

from bill_detail_automation.settings import BASE_DIR


class Main(View):
	return_url = None

	def post(self, request):
		FilePath = request.FILES['file']
		response=self.main_fun(FilePath,1, 1)
		return response
		# return render(request, 'index.html')

	def listToString(self, s):
		# initialize an empty string
		str1 = ""

		# traverse in the string
		for ele in s:
			str1 += str(ele) + ','

		# return string
		return str1

	def AVVNL_JDVVNL_helper(self, SheetName, driver):
		if SheetName == 'AVVNL' or SheetName == 'avvnl':
			for i in range(2):
				try:
					myElem = driver.find_element(By.XPATH,
												 '/html/body/div[1]/section/div/div/div[3]/div[2]/div[69]/div/a')
					print("main Page is ready!")
					break
				except:
					print("main page took too much time!")
					time.sleep(3)
			AVVNL_button = driver.find_element(By.XPATH,
											   '/html/body/div[1]/section/div/div/div[3]/div[2]/div[69]/div/a')
			# AVVNL_button.click()
			driver.execute_script('arguments[0].click()', AVVNL_button)
		elif SheetName == 'JDVVNL' or SheetName == 'JDVVNL':
			for i in range(2):
				try:
					myElem = driver.find_element(By.XPATH,
												 '/html/body/div[1]/section/div/div/div[3]/div[2]/div[68]/div/a')
					print("main Page is ready!")
					break
				except:
					print("main page took too much time!")
					time.sleep(3)
			JDVVNL_button = driver.find_element(By.XPATH,
												'/html/body/div[1]/section/div/div/div[3]/div[2]/div[68]/div/a')
			# JDVVNL_button.click()
			driver.execute_script('arguments[0].click()', JDVVNL_button)

	def main_fun(self,FilePath,MMT, SPT):
		print(FilePath)
		xls = pd.ExcelFile(FilePath)
		print(xls.sheet_names)
		# driver = webdriver.Chrome(resource_path('.\drivers\chromedriver'))
		# driver = webdriver.Chrome('.\drivers\chromedriver')
		driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
		driver.maximize_window()
		cols = ["Kno", "Receipt No", "Amount", "Receipt Date", "AVVNL/JDVVNL"]
		rows = []
		missing_list = []
		for SheetName in xls.sheet_names:
			df = pd.read_excel(FilePath, sheet_name=SheetName, index_col=None, usecols="A")
			knumbers = []
			for column in df.columns:
				knumbers = df[column].tolist()
			print(knumbers)
			driver.get('https://jansoochna.rajasthan.gov.in/Services')
			time.sleep(MMT)
			delay = 30
			self.AVVNL_JDVVNL_helper(SheetName, driver)
			a = 1

			for kno in knumbers:
				driver.refresh
				tempList = []
				MyInput = int(kno)
				while True:
					try:
						myElem = driver.find_element(By.XPATH, '//*[@id="Enter_your_K_number"]')
						print("search Page is ready!")
						break
					except:
						print("search page took too much time!")
						driver.refresh
						time.sleep(SPT)
				box = driver.find_element(By.XPATH, '//*[@id="Enter_your_K_number"]')
				box.send_keys(MyInput)
				searchButton = driver.find_element(By.NAME, 'खोजें')
				# searchButton.click()
				driver.execute_script('arguments[0].click()', searchButton)
				time.sleep(SPT)
				th_count = driver.find_elements(By.TAG_NAME, 'th')
				print("th count is:" + str(len(th_count)))
				if len(th_count) < 4:
					print("skipped kno:" + str(kno))
					driver.get('https://jansoochna.rajasthan.gov.in/Services')
					time.sleep(MMT)
					self.AVVNL_JDVVNL_helper(SheetName, driver)
					tempList.append(kno)
					tempList.append("No Data Found on Website!!")
					missing_list.append(tempList)
					a = a + 1
					continue

				if SheetName == 'AVVNL' or SheetName == 'avvnl':
					while True:
						try:
							ReceiptNo = driver.find_element(By.XPATH, '//*[@id="tblResult_0"]/tbody/tr[1]/td[6]').text
							Scrap_kno = int(driver.find_element(By.XPATH, '//*[@id="tblResult_0"]/tbody/tr[1]/td[1]').text)
							ReceiptDate = driver.find_element(By.XPATH, '//*[@id="tblResult_0"]/tbody/tr[1]/td[7]').text

							Amount = driver.find_element(By.XPATH, '//*[@id="tblResult_0"]/tbody/tr[1]/td[9]').text
							if Scrap_kno == MyInput:
								tempList.append(Scrap_kno)
								tempList.append(ReceiptNo)
								tempList.append(Amount)
								tempList.append(ReceiptDate)
							else:
								raise TypeError("do it once again")

							print(MyInput, "data added")
							break
						except Exception as e:
							print("table took too much time!")
							time.sleep(SPT)
							box.clear()
							box = driver.find_element(By.XPATH, '//*[@id="Enter_your_K_number"]')
							box.send_keys(MyInput)
							searchButton = driver.find_element(By.NAME, 'खोजें')
							driver.execute_script('arguments[0].click()', searchButton)
							time.sleep(SPT)

				elif SheetName == 'JDVVNL' or SheetName == 'jdvvnl':
					while True:
						try:
							ReceiptNo = driver.find_element(By.XPATH, '//*[@id="tblResult_0"]/tbody/tr[1]/td[6]').text
							Scrap_kno = int(driver.find_element(By.XPATH, '//*[@id="tblResult_0"]/tbody/tr[1]/td[1]').text)
							ReceiptDate = driver.find_element(By.XPATH, '//*[@id="tblResult_0"]/tbody/tr[1]/td[13]').text

							Amount = driver.find_element(By.XPATH, '//*[@id="tblResult_0"]/tbody/tr[1]/td[12]').text
							if Scrap_kno == MyInput:
								tempList.append(Scrap_kno)
								tempList.append(ReceiptNo)
								tempList.append(Amount)
								tempList.append(ReceiptDate)
							else:
								raise TypeError("do it once again")

							print(MyInput, "data added")
							break
						except Exception as e:
							print("table took too much time!")
							time.sleep(SPT)
							box.clear()
							box = driver.find_element(By.XPATH, '//*[@id="Enter_your_K_number"]')
							box.send_keys(MyInput)
							searchButton = driver.find_element(By.NAME, 'खोजें')
							driver.execute_script('arguments[0].click()', searchButton)
							time.sleep(SPT)

				tempList.append(SheetName)
				rows.append(tempList)
				print("output is here")
				print(a, tempList)
				a = a + 1
				box.clear()
		for myRow in missing_list:
			rows.append(myRow)
		extension = ".xlsx"
		Filename ="Kishan"+ time.strftime("%Y%m%d-%H%M%S") + extension
		driver.quit()
		df = pd.DataFrame(rows,columns=cols)
		# writer = pd.ExcelWriter(Filename, engine='xlsxwriter')
		#
		# df.to_excel(writer, sheet_name='Sheet1', index=False)
		#
		# writer.save()

		response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
		response['Content-Disposition'] = 'attachment; filename="'+str(Filename)+'.xlsx"'
		df.to_excel(response)
		return response


	# def download_file(self,filename):
	# 	# Define the full file path
	# 	BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
	# 	filepath = str(BASE_DIR) + '/' + filename
	# 	# Open the file for reading content
	# 	path = open(filepath, 'r')
	# 	# Set the mime type
	# 	mime_type, _ = mimetypes.guess_type(filepath)
	# 	# Set the return value of the HttpResponse
	# 	response = HttpResponse(path, content_type=mime_type)
	# 	# Set the HTTP header for sending to browser
	# 	response['Content-Disposition'] = "attachment; filename=%s" % filename
	# 	# Return the response value
	# 	return response