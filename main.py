import pandas as pd
import datetime
import sys


class excelExport():
    def read_file(self) -> pd.DataFrame:
        ''' Reads excel file and return pandas dataframe '''

        df = pd.read_excel('input_data.xlsx')
        return df

    def get_tax_rate(self, state) -> dict:
        '''Gets tax rate based on state '''

        state_rates = {'IL': .0251, 'TN': .01766, }
        return state_rates[state]

    def write_new_file(self, aggregated_data_frame: pd.DataFrame, report_date: str):
        ''' Writes Pandas dataframe data to an excel sheet '''

        aggregated_data_frame.to_excel(f'aggregated_report-{report_date}.xlsx', index=False)
        return

    def validateExcelData(self, df: pd.DataFrame) -> dict:
        '''
        Checks if the data in the dataframe is malformed.
        '''
        malformData = {}
        vinIndexList = self.checkVIN(inputDF=df)
        effectiveDateList = self.checkDate(inputDF=df, column='Effective Date')
        expirationDateList = self.checkDate(
            inputDF=df, column='Expiration Date')
        annualGwpIndex = self.checkAnnualGWP(inputDF=df)

        while len(vinIndexList) > 0 or len(effectiveDateList) > 0 or len(expirationDateList) > 0 or len(annualGwpIndex) > 0:
            print('*'*15 + 'WARNING' + '*'*15)
            if len(vinIndexList) > 0:
                print('Malform data in column \'VIN\' index = %s' % vinIndexList)
                malformData['VIN'] = vinIndexList
                vinIndexList = []
            elif len(effectiveDateList) > 0:
                print('Malform data in column \'Effective Date\' index = %s' %
                      effectiveDateList)
                malformData['Effective Date'] = effectiveDateList
                effectiveDateList = []
            elif len(effectiveDateList) > 0:
                print('Malform data in column \'Expiration Date\' index = %s' %
                      expirationDateList)
                malformData['Expiration Date'] = effectiveDateList
                effectiveDateList = []
            elif len(annualGwpIndex) > 0:
                print('Malform data in column \'Annual GWP\' index = %s' %
                      annualGwpIndex)
                malformData['Annual GWP'] = annualGwpIndex
                annualGwpIndex = []

        return malformData

    def checkVIN(self, inputDF: pd.DataFrame) -> list:
        '''
        Checks if VIN data contains the valid 17 char limit
        '''
        try:
            vinList = inputDF['VIN'].tolist()
        except KeyError:
            print('Dataframe doesnt have VIN column')
            return

        malformVinRow = []
        for i in range(0, len(vinList)):
            if len(vinList[i].strip()) != 17:
                malformVinRow.append(i)
        return malformVinRow

    def checkDate(self, inputDF: pd.DataFrame, column: str) -> list:
        ''' Checks if the date string is in a valid date format '''
        try:
            inputDF[column]
        except KeyError:
            print('Dataframe doesnt have Date column')
            return

        expirationDateList = inputDF[column].map(type).tolist()
        malformeExpirationDate = []

        for i in range(0, len(expirationDateList)):
            if expirationDateList[i] is pd.Timestamp or expirationDateList[i] is datetime.datetime:
                continue
            else:
                malformeExpirationDate.append(i)
        return malformeExpirationDate

    def checkAnnualGWP(self, inputDF: pd.DataFrame) -> list:
        ''' checks if the Annual GWP is an intiger '''

        try:
            annualGwpList = inputDF['Annual GWP'].tolist()
        except KeyError:
            print('Dataframe doesnt have Annual GWP column')
            return

        malformAnnualGwpRow = []
        for i in range(0, len(annualGwpList)):
            if type(annualGwpList[i]) != int:
                malformAnnualGwpRow.append(i)
        return malformAnnualGwpRow

    def checkRDateWithin(self, inputDF: pd.DataFrame, reportedDate: datetime.datetime):
        '''
        Checks if report date is NOT between effective date 
        and Expiration date. Returns a list of index
        '''
        try:
            inputDF['Expiration Date']
            inputDF['Effective Date']
        except KeyError:
            print('Dataframe doesnt have a pair Date columns to compare')
            return

        reportDateWithinDateWindowList = []
        for ind in inputDF.index:
            if inputDF['Expiration Date'][ind] < reportedDate or inputDF['Effective Date'][ind] > reportedDate:
                reportDateWithinDateWindowList.append(ind)
        if len(reportDateWithinDateWindowList):
            print('*'*15 + 'WARNING' + '*'*15)
            print('Report day is outside the Effective/Expiration Date for index %s' %
                  reportDateWithinDateWindowList)

    def addcolumns(self, inputDF: pd.DataFrame, reportedDate: datetime.datetime) -> pd.DataFrame:
        '''
        Adds additional columns to the input df and performs calulation to insert data
        into new columns.
        Additional columns are
        - Pro-rata Gross Written Premium
        - Earned premium
        - Unearned premium
        - Taxes 
        '''
        proRatGwpList = []
        earnedPremiumList = []
        unearnedPremiumList = []
        taxAmountList = []

        for ind in inputDF.index:
            dailyGWP = inputDF['Annual GWP'][ind]/365
            effectiveDays = (inputDF['Expiration Date']
                             [ind]-inputDF['Effective Date'][ind]).days

            # Calculation on getting pro-rata GWP
            ProGwp = dailyGWP*effectiveDays
            proRatGwpList.append(ProGwp)

            # Calculating earned premium
            effectedToReportedDate = (
                reportedDate - inputDF['Effective Date'][ind]).days
            earnedPremium = dailyGWP*effectedToReportedDate
            earnedPremiumList.append(earnedPremium)

            # Calculate unearned premium
            repotedToExpirationDate = (
                inputDF['Expiration Date'][ind]-reportedDate).days
            unearnedPremium = dailyGWP*repotedToExpirationDate
            unearnedPremiumList.append(unearnedPremium)

            # Calculate Tax amount, this is done by earnedPremium * tax, not sure if this is correct
            taxAmount = earnedPremium * \
                self.get_tax_rate(state=inputDF['State'][ind])
            taxAmountList.append(taxAmount)

        # Appends columns with the calculated data
        inputDF['Prorata GWP'] = proRatGwpList
        inputDF['Earned Premium'] = earnedPremiumList
        inputDF['Unearned Premium'] = unearnedPremiumList
        inputDF['Tax Amount'] = taxAmountList

        return inputDF

    def createOutputDF(self, inputDf: pd.DataFrame, rDate) -> pd.DataFrame:
        '''
        Aggregate by Company Name to count number of vehicles/VINs and sum 
        vehicle pro-rata GWP, earned premium, unearned premium, and taxes
        '''
        vinList = []
        totalAnnualGwpList = []
        totalProRataGwpList = []
        totalEarnedPremiumList = []
        totalUnearnedPremiumList = []
        totalTaxAmountList = []
        outputDF = pd.DataFrame(columns=['Company Name', 'Report Date', 'Total VIN Count', 'Total Annual GWP',
                                         'Total Pro-Rata GWP', 'Total Earned Premium', 'Total Unearned Premium', 'Total Taxes'])
        # Write all unque Company name
        companyNameList = inputDf['Company Name'].unique()
        outputDF['Company Name'] = companyNameList

        for name in companyNameList:
            # add VIN count
            filteredDf = inputDf[inputDf['Company Name'] == name]
            vinList.append(filteredDf['VIN'].nunique())

            # Calculate total annual GWP
            totalAnnualGwpList.append(filteredDf['Annual GWP'].sum())

            # Getting the total Pro-rata GWP
            totalProRataGwpList.append(filteredDf['Prorata GWP'].sum())

            # Getting the total Earned Premium
            totalEarnedPremiumList.append(filteredDf['Earned Premium'].sum())

            # Getting the total Unearned Premium
            totalUnearnedPremiumList.append(
                filteredDf['Unearned Premium'].sum())

            # Getting the total Tax amount
            totalTaxAmountList.append(filteredDf['Tax Amount'].sum())

        # Ataching the new rows to the columns
        outputDF['Report Date'] = rDate
        outputDF['Total VIN Count'] = vinList
        outputDF['Total Annual GWP'] = totalAnnualGwpList
        outputDF['Total Pro-Rata GWP'] = totalProRataGwpList
        outputDF['Total Earned Premium'] = totalEarnedPremiumList
        outputDF['Total Unearned Premium'] = totalUnearnedPremiumList
        outputDF['Total Taxes'] = totalTaxAmountList

        return outputDF


if __name__ == '__main__':
    rDate = '2022-08-01'
    excelObject = excelExport()

    df = excelObject.read_file()
    malformDict = excelObject.validateExcelData(df=df)

    # Removes the rows that contain the malform data ex: 100 != 1o0 or 2021-10-O1 != 2021-10-01
    if len(malformDict) > 0:
        print('-'*40)
        print('*'*5+'Removing rows that contain malform data' + '*'*5)
        print('-'*40)
        for key, value in malformDict.items():
            df = df.drop(malformDict[key])
        df.reset_index(drop=True, inplace=True)

    excelObject.checkRDateWithin(
        inputDF=df, reportedDate=datetime.datetime.strptime(rDate, '%Y-%m-%d'))

    # Append column with the calculated data
    df = excelObject.addcolumns(
        inputDF=df, reportedDate=datetime.datetime.strptime(rDate, '%Y-%m-%d'))
    df = excelObject.createOutputDF(df, rDate)

    excelObject.write_new_file(df, rDate)
    print('Creating a new excel file with Aaggregate data')
