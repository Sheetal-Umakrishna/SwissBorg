import openpyxl
import XLUtils
from selenium import webdriver
import unittest
from time import sleep

class CalculatorTest(unittest.TestCase):
    def test_Chrome(self):

        # Installing Driver and Launching the Website
        self.driver=webdriver.Chrome(executable_path="C:/Users/Sheetal/Desktop/Testing/chromedriver.exe")
        self.driver.get("http://qainterview.pythonanywhere.com/")

        # Reading and writing from the .xlsx file
        path="C:/Users/Sheetal/Desktop/myexcel.xlsx"
        rows=XLUtils.getRowCount(path,"Sheet1")

        #looping through the excel file
        for r in range(1,rows+1):

            #Reading the excel file
            data=XLUtils.readData(path,"Sheet1",r,1)

            #sending input
            self.driver.find_element_by_name("number").send_keys(data)

            #clicking the Calculate button
            self.driver.find_element_by_id("getFactorial").click()

            result = self.driver.find_element_by_xpath("//p[@id='resultDiv']")
            result2 = result.is_displayed()
            result3 = result.text

            # Verifying the result
            if result2==True:
                print("test is passed")

                #writing into the excel file
                XLUtils.writeData(path,"Sheet1",r,2,result3)

            else:
                print("test failed")
                XLUtils.writeData(path,"Sheet1",r,2,"testfailed")
    
            self.driver.find_element_by_name("number").clear()

        #Verifying the correctness of the calcultor for integer number in range(10,100)
        for r in range(2,rows):
            self.assertEqual(XLUtils.readData(path,"Sheet1",r,2),XLUtils.readData(path,"Sheet1",r,3))

if __name__ == "__main__":
   unittest.main()
    








   
        


    



