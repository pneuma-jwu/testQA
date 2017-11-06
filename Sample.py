# -*- coding: utf-8 -*-
from selenium import webdriver
from datetime import datetime
import time
import random
import string
import win32com.client
import sys
if sys.version_info[0] < 3:
    #raise "Must be using Python 3"
	from Tkinter import *
	import Tkinter as tk
else:
	from tkinter import *
	import tkinter as tk
import logging
import os

# create logger
# logger = logging.getLogger(__name__)
#logger = logging.getLogger(__file__)
#script_name = os.path.basename('C:\QAAutomation\Cliqstudios\TEST\Sample Order with New Accout-auto click start test.py')
script_name = "QA_TASK"
logger = logging.getLogger(script_name);
logger.setLevel(logging.DEBUG)

# create console handler and set level to debug
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)

# create formatter
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# add formatter to ch
ch.setFormatter(formatter)

# add ch to logger
logger.addHandler(ch)

'''
#Logging tests
# 'application' code
logger.debug('debug message')
logger.info('info message')
logger.warn('warn message')
logger.error('error message')
logger.critical('critical message')
'''
now_date = str(datetime.now().strftime('%m%d'))
#screenshots_folder = r"D:\Pneuma Project\AutoTest\Cliq_webdriver\Test\\"
screenshots_folder = r"D:\Jason\Documents\QAAutomation\Screenshots\\" 


def callback():
	logger.debug("Entered callback()")
	#chromedriver = executable_path=r'C:\chromedriver\chromedriver.exe'
	chromedriver  = executable_path=r'D:\Install\chromedriver\chromedriver.exe'
	driver = webdriver.Chrome(chromedriver)
	driver.maximize_window()
	driver.implicitly_wait(10)
	#driver.get('https://equus:quality2015@equus.cliqstudios.com/)
	driver.get('http://172.30.9.19')
	shell = win32com.client.Dispatch("WScript.Shell")   
	
	logger.info("shell.Sendkeys(\"equus\")")
	shell.Sendkeys("equus")  
	time.sleep(2)
	logger.info("shell.Sendkeys(\"{TAB}\")")
	shell.Sendkeys("{TAB}")
	time.sleep(2)
	logger.info("shell.Sendkeys(\"quality2015\")")
	shell.Sendkeys("quality2015") 
	time.sleep(2)
	logger.info("shell.Sendkeys(\"{ENTER}\")")
	shell.Sendkeys("{ENTER}")
	time.sleep(2)
	
	count = 1
	n1 = int(sn.get())
	while count <= int(fq.get()):
		logger.info("Perform #%d time tests." % (count))
		time.sleep(3)
		#Click Get Free Sample
		logger.info("Click Get Free Sample")
		#driver.find_element_by_xpath('//*[@id="preheader"]/div/div[2]/div[1]/a').click()
		driver.find_element_by_id('get_free_sample').click()
		time.sleep(1)
		#Click Add to Cart
		logger.info("Click Add to Cart")
		driver.find_element_by_id('AddToCartCS').click()
		time.sleep(1)
		#Click Proceed to Checkout
		logger.info("Click Proceed to Checkout")
		driver.find_element_by_id('img_proceed_to_checkout').click()
		time.sleep(1)
		#Input First Name
		logger.info("Input First Name")
		driver.find_element_by_id('first_name').send_keys(fn.get() + '%03d'%n1)
		#Input Last Name
		logger.info("Input Last Name")
		driver.find_element_by_id('last_name').send_keys(ln.get())
		#Input Email
		logger.info("Input Email")
		driver.find_element_by_id('u_name_new').send_keys(fn.get() + now_date + '%03d'%n1 + '@cliqTEST.com')
		#driver.find_element_by_id('u_name_new').send_keys(fn.get() + now_date + '%03d'%n1 + '@equus.cliqstudios.com')
		#Input Phone Number
		logger.info("Input Phone Number")
		driver.find_element_by_id('phone_number_p1').send_keys(str(random.randint(100,999)))
		driver.find_element_by_id('phone_number_p2').send_keys('555')
		driver.find_element_by_id('phone_number_p3').send_keys(str(random.randint(1000,9999)))
		#Input Password
		logger.info("Input Password")
		driver.find_element_by_id('user_password_new').send_keys('123456')
		driver.find_element_by_id('user_password_confirm').send_keys('123456')
		#Input Zipcode
		logger.info("Input Zipcode")
		driver.find_element_by_id('postal_code').send_keys(random.choice(['55343','10004','50012','43215']))
		#Click create button
		logger.info("Click create button")
		driver.find_element_by_id('btnNewAccountCheck').click()
		time.sleep(1)
		#Input Address 1  
		logger.info("Input Address 1  ")
		driver.find_element_by_name('address_2').send_keys(random.choice(['123 Main Steet','379 Timbercrest Road','626 Mayfield Dr','58 Medical Center Drive']))
		#Input Address 2
		logger.info("Input Address 2  ")
		driver.find_element_by_name('apt_2').send_keys(random.choice(['Apt#150','','Roof']))
		#Input City
		logger.info("Input City  ")
		driver.find_element_by_name('city_2').send_keys(random.choice(['New York','Chicago','Las Vegas']))
		#Choose State
		logger.info("Choose State  ")
		#driver.find_element_by_xpath('//*[@name="state_code_2"]/option[' + str(random.randint(0,50)) + ']').click()
		driver.find_element_by_xpath('//*[@name="state_code_2"]/option[' + str(random.randint(0,50)) + ']').click()
		#Input Alternate Phone
		logger.info("Input Alternate Phone  ")
		driver.find_element_by_name('p1_2').send_keys(str(random.randint(100,999)))
		driver.find_element_by_name('p2_2').send_keys('555')
		driver.find_element_by_name('p3_2').send_keys(str(random.randint(1000,9999)))
		#save address as default
		#logger.info("save address as default  ")
		#driver.find_element_by_id('imgBtnNext').click()
		#Click Next
		logger.info("Click Next  ")
		driver.find_element_by_id('imgBtnNext').click()
		time.sleep(1)
		#Click Submit Order
		logger.info("Click Submit Order ")
		driver.find_element_by_xpath('//*[@id="imgBtnSubmit"]').click()
		time.sleep(2)
		driver.execute_script("window.scrollTo(0, 590)")
		time.sleep(2)
		driver.get_screenshot_as_file( screenshots_folder + 'SampleOrder' + str(n1) +'.png')
		driver.find_element_by_xpath('//*[@id="hello_user"]/span').click()
		count = count + 1
		n1 = n1 + 1
		time.sleep(2)
	logger.debug("Quiting webdriver ")
	driver.quit()
	logger.debug("Quiting callback()")
	
def id_generator(size=3, chars=string.ascii_uppercase + string.digits):
	return '-'.join(random.choice(chars) for _ in range(size))	

def limitSizeSN(*args):
    value = SnValue.get()
    if len(value) > 3: SnValue.set(value[:3])


window = tk.Tk()    
window.title('QATEST_Sample Order')

tk.Label(window, text = 'Please enter your test data', font=('Calibri', 11)).grid(row=0, sticky=W,columnspan=2)

tk.Label(window, text='Serial No:',font=('Calibri',11)).grid(row=1, column=0, sticky=W)
SnValue = StringVar(value='123')
SnValue.trace('w', limitSizeSN)

sn = tk.Entry(window, textvariable=SnValue, font=('Calibri',11), show=None)
sn.grid(row=1, column=1, sticky=W,padx=5)
sn.focus_set()

tk.Label(window, text = 'Frequency:',font=('Calibri',11)).grid(row=2, column=0, sticky=W)
fq = tk.StringVar(value='1')
fq = tk.Entry(window, textvariable=fq, font=('Calibri',11), show=None)
fq.grid(row=2, column=1, sticky=W,padx=5)

tk.Label(window, text='First name:',font=('Calibri',11)).grid(row=3, column=0, sticky=W)
fn = tk.StringVar(value='Anthony' + id_generator() )
fn = tk.Entry(window, textvariable=fn, font=('Calibri',11), show=None)
fn.grid(row=3, column=1, sticky=W,padx=5)

tk.Label(window, text='Last name:',font=('Calibri',11)).grid(row=4, column=0, sticky=W)
ln = tk.StringVar(value='Hoang'  + id_generator() )
ln = tk.Entry(window, textvariable=ln, font=('Calibri',11), show=None)
ln.grid(row=4, column=1, sticky=W,padx=5)

tk.Label(window, text='URL:',font=('Calibri',11)).grid(row=5, column=0, sticky=W)
u1 = tk.StringVar()
u1 = tk.Entry(window, textvariable=u1, font=('Calibri',11), show=None, state='disable')
u1.grid(row=5, column=1, sticky=W,padx=5)

b1 = Button(window, text="Start Test", font=('Calibri',11), width=10, command=callback)
b1.grid(row=6, column=0,sticky=E, padx=5, pady=5)
b2 = Button(window, text="End Test", font=('Calibri',11), width=10, command=window.destroy)
b2.grid(row=6, column=1, sticky=W)

var = tk.StringVar()
d = tk.Label(window, textvariable=var, fg='red')
d.grid(row=7, column=0, sticky=W,columnspan=2)
logger.info("QA task starting...")
b1.invoke()
b2.invoke()
logger.info("Normal End.")



#thefrawindow = 
window.mainloop()


