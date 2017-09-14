# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import os
import sys
import re
import requests
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver

cik = input('Please input the TICKER SYMBOL or CIK of the corporation(e.g. Nuan), CASE-INSENSITIVE: ')
year = input('Please input the year in which the report was released(e.g. : 2016): ')
types = input('Please input the type of the report(10-Q or 10-K), CASE-INSENSITIVE: ')
if types in ['10-Q', '10-q']:
	quart = input('Please input the number of QUARTERS(1 or 2 or 3): ')

def href(url):	# This function is used to return the link to be visited.
	driver.get(url)
	String = driver.page_source
	soup_0 = BeautifulSoup(String, 'lxml')
	sml_0 = soup_0.find(name = 'table', attrs = {'class': 'tableFile2', 'summary': 'Results'})
	trs_list_0 = sml_0.find_all(name = 'tr')
	del trs_list_0[0]
	trs_list = []
	for das in trs_list_0:
		dich = das.find(name = 'td', attrs = {'class': 'small'})
		if dich.find_all(name = 'b') == []:
			trs_list.append(das)
	part = trs_list[0].find_all(name = 'td')[1].contents[0]['href']
	for das in trs_list[0].find_all(name = 'td'):
		attr = list(das.attrs.keys())
		if (not 'nowrap' in attr and not 'class' in attr):
			date = das.string
			break
	return (part,date)

def checktime(year, url):	#	This function is used to examine whether the year input is valid or not.
	driver.get(url)
	String = driver.page_source
	soup_0 = BeautifulSoup(String, 'lxml')
	sml_0 = soup_0.find(name = 'table', attrs = {'class': 'tableFile2', 'summary': 'Results'})
	trs_list_0 = sml_0.find_all(name = 'tr')
	del trs_list_0[0]
	trs_list = []
	date_list = []
	for das in trs_list_0:
		doc_flag = 0
		dich = das.find(name = 'td', attrs = {'class': 'small'})
		if dich.find_all(name = 'b') == []:
			trs_list.append(das)
			doc_flag = 1
		if doc_flag == 1:
			ich = das.find_all(name = 'td')
			date_list.append(ich[3].string)
	year_list = []
	for item in date_list:
		if year in item:
			year_list.append(item)
	if year_list != []:
		return 1
	else:
		return 0

def quarter(year, quart, url):	#	This function is used to return the link to be visited.
	driver.get(url)
	String = driver.page_source
	soup_0 = BeautifulSoup(String, 'lxml')
	sml_0 = soup_0.find(name = 'table', attrs = {'class': 'tableFile2', 'summary': 'Results'})
	trs_list_0 = sml_0.find_all(name = 'tr')
	del trs_list_0[0]
	trs_list = []
	date_list = []
	for das in trs_list_0:
		doc_flag = 0
		dich = das.find(name = 'td', attrs = {'class': 'small'})
		if dich.find_all(name = 'b') == []:
			trs_list.append(das)
			doc_flag = 1
		if doc_flag == 1:
			ich = das.find_all(name = 'td')
			date_list.append(ich[3].string)
	year_list = []
	for item in date_list:
		if year in item:
			year_list.append(item)
	year_list.reverse()
	if len(year_list) < int(quart):
		print('The report is not available at present. Please check the quarter number that you\'ve input')
		print('The program will be ended automatically after 6 Mississippi')
		time.sleep(6)
		sys.exit(0)
		part = 0
	else:
		dim = year_list[int(quart) - 1]
		exact = date_list.index(dim)
		part = trs_list[exact].find_all(name = 'td')[1].contents[0]['href']
	return part
	
def end_process(para):
	if para == 0:
		time.sleep(6)
		sys.exit(0)

print('**********************Phase I : Searching for the report, not for very long**********************')	#	Search for the report in EDGAR SYSTEM.
if types in ['10-K', '10-Q', '10-k', '10-q']:
	url_0 = 'https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=' + cik + '&type=' + types + '&dateb=' + year + '1231' + '&owner=exclude&count=100'
else:
	print('The report type that you\'ve input is invalid and the program will be ended automatically after 6 Mississippi')
	end_process(0)
if types in ['10-Q', '10-q']:
	if not quart in ['1', '2', '3']:
		print('The quarter number that you\'ve input is invalid and the program will be ended automatically after 6 Mississippi')
		end_process(0)
driver = webdriver.Chrome()
driver.get(url_0)
str_2 = driver.page_source
soup_1 = BeautifulSoup(str_2, 'lxml')
if 'No matching Ticker Symbol' in str_2:
	print('The TICKER SYMBOL or CIK that you input is wrong')
	driver.quit()
	print('The program will be ended automatically after 6 Mississippi')
	end_process(0)
else:
	sml_1 = soup_1.find(name = 'table', attrs = {'class': 'tableFile2', 'summary': 'Results'})
	if len(sml_1.contents) == 1:
		print('The report type that you input is wrong')
		driver.quit()
		print('The program will be ended automatically after 6 Mississippi')
		end_process(0)
	else:
		check_sym = checktime(year, url_0)
		if check_sym == 0:
			print('The report is not available at present. Please check the year that you\'ve input')
			print('The program will be ended automatically after 6 Mississippi')
			end_process(0)
		else:
			if types in ['10-K', '10-k']:
				shilly = href(url_0)
				if shilly[1][5:7] in ['01', '02', '03', '04', '05', '06']:
					url_0 = re.sub(year, str(int(year) + 1), url_0)
					shilly = href(url_0)
					url_1 = 'https://www.sec.gov' + shilly[0]
				else:
					url_1 = 'https://www.sec.gov' + shilly[0]
			else:
				shilly = quarter(year, quart, url_0)
				url_1 = 'https://www.sec.gov' + shilly
			driver.get(url_1)
			string = driver.page_source
			soup_x = BeautifulSoup(string, 'lxml')
			url_2 = 'https://www.sec.gov' + soup_x.find(name = 'table', attrs = {'class': 'tableFile', 'summary': 'Document Format Files'}).find_all(name = 'tr')[1].find_all(name = 'td')[2].find(name = 'a')['href']
			if url_2[len(url_2) - 3: len(url_2)] == 'txt':
				print('The report is too old to be transformed into xlsx')
				driver,quit()
				print('The program will be ended automatically after 6 Mississippi')
				end_process(0)
			else:
				driver.get(url_2)
				str_1 = driver.page_source
				driver.quit()
			
			
		