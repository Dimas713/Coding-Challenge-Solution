import unittest
import sys
import os
import pandas as pd

# Allows us to import excelExport object
sys.path.insert(1, os.path.join(sys.path[0], '..'))

from main import excelExport

class BasicTests(unittest.TestCase):
    def test_checkVIN(self):
        '''Testing checkVIN function reaction towards empty dataframes'''
        excelObject = excelExport()
        testDf = pd.DataFrame()
        
        excelObject.checkVIN(testDf)
        return

    def test_checkVINMalform(self):
        '''Testing checkVIN function reaction towards malfomred dataframes'''
        excelObject = excelExport()
        testDf = pd.DataFrame()

        testDf['VIN'] = ['1f8s2s8dS','0000000']

        vinList = excelObject.checkVIN(testDf)

        self.assertTrue(vinList)
        return
    
    def test_checkDate(self):
        '''Testing checkDate function reaction towards empty dataframes'''
        excelObject = excelExport()
        testDf = pd.DataFrame()

        excelObject.checkDate(inputDF= testDf, column= 'Expiration Date')
        return

    def test_checkDateMalform(self):
        '''Testing test_checkDateMalform function reaction towards malfomred dataframes'''
        excelObject = excelExport()
        testDf = pd.DataFrame()
        testDf['Expiration Date'] = ['10/12/2022','10/ff/2022']
        dateList = excelObject.checkDate(inputDF= testDf, column= 'Expiration Date')

        self.assertTrue(dateList)
        return

    def test_checkAnnualGWP(self):
        '''Testing checkAnnualGWP function reaction towards empty dataframes'''
        excelObject = excelExport()
        testDf = pd.DataFrame()
        
        excelObject.checkAnnualGWP(testDf)
        return
    
    def test_checkAnnualGWPMalform(self):
        '''Testing checkAnnualGWP function reaction towards malfomred dataframes'''
        excelObject = excelExport()
        testDf = pd.DataFrame()

        testDf['Annual GWP'] = ['1f8s2s8dS','4rr444']

        annualGwp = excelObject.checkAnnualGWP(testDf)

        self.assertTrue(annualGwp)
        return
    def test_checkRDateWithin(self):
        '''Testing test_checkRDateWithin function reaction towards empty dataframes'''
        excelObject = excelExport()
        testDf = pd.DataFrame()

        excelObject.checkRDateWithin(inputDF= testDf, reportedDate= '10-01-2022')
        return

    def test_checkRDateWithinMalform(self):
        '''Testing test_checkRDateWithinMalform function reaction towards date data that are malformed'''
        excelObject = excelExport()
        testDf = pd.DataFrame()
        testDf['Expiration Date'] = ['02/13/2022']
        testDf['Effective Date'] = ['02/13/2022']
        excelObject.checkRDateWithin(inputDF= testDf, reportedDate= '10-01-2022')
        return


if __name__ == '__main__':
    unittest.main()