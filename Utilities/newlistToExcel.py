
# Copyright 2025 Adobe. All rights reserved.
# This file is licensed to you under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License. You may obtain a copy
# of the License at http://www.apache.org/licenses/LICENSE-2.0

# Unless required by applicable law or agreed to in writing, software distributed under
# the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
# OF ANY KIND, either express or implied. See the License for the specific language
# governing permissions and limitations under the License.


def newgenerateExcelFromListJP_KO_CN_TW(excel,filename,JPintlExcelList,JPresults,KOintlExcelList,KOresults,zhCNintlExcelList,zhCNresults,zhTWintlExcelList,zhTWresults):
    print(JPresults)
    print(KOresults)
    print(zhCNresults)
    import xlsxwriter
    workbook = xlsxwriter.Workbook(filename)

    sheet1 = workbook.add_worksheet('String Rendering')
    sheet1.set_column(0,len(excel[0]),20)
   
    # sheet2.autofit()
    bold = workbook.add_format({'bold': True})
    passFormat=workbook.add_format({'bold': True, 'font_color': 'green'})
    failFormat=workbook.add_format({'bold': True, 'font_color': 'red'})
    sheet1.write(0, 0, 'Locale',bold)
    sheet1.write(0, 1, 'Test String',bold)
    sheet1.write(0, 2, 'Rendered String',bold)
    sheet1.write(0, 3, 'Actual String',bold)
    sheet1.write(0, 4, 'Result',bold)
    sheet1.write(0,5,'Screen Location',bold) #Not there in manualAdded-Mail-report
        
    for row in range(len(excel)):
        for col in range(len(excel[row])):
            print(row,col,excel[row][col],sep=" - ")
            if col ==5:
                modi=excel[row][col]
                print(modi)
                sheet1.write_url(row+1,col,modi,string="Open Screen")
            else:
                if(str(excel[row][col]) == "Pass"):
                    sheet1.write(row+1,col,str(excel[row][col]),passFormat)
                elif(str(excel[row][col]) == "Fail"):
                    sheet1.write(row+1,col,str(excel[row][col]),failFormat)
                else:
                    sheet1.write(row+1,col,str(excel[row][col]))
    sampleLink="https://jira.corp.adobe.com/browse/{}"
    ## Now Sheet 2 for JP Intl
    if([] not in JPresults):
        sheet2= workbook.add_worksheet('INTL JP Test Cases')
        sheet2.write(0,0,"Test Case Description",bold)
        # sheet2.write(0,1,"Result",bold)
        sheet2.write(0,1,"Jira Links",bold)
        currentCol=2
        for values in range(len(JPresults[0])):
            sheet2.write(0,currentCol,"Target Dialog "+str(values+1),bold)
            currentCol+=1
        #JPresults=[['Pass','B-Pass'], ['Pass','B-Fail'], ['Pass','B-Pass'], ['Pass','B-Fail'], ['Pass','B-Pass']]
        for row in range(len(JPintlExcelList)):
            sheet2.write(row+1,0,JPintlExcelList[row])
     
            sheet2.write_url(row+1,1,sampleLink.format(JPintlExcelList[row].split(':')[0].strip().replace(' ','-')),string="Open Jira Link")
            for resultValue in range(len(JPresults[row])):
                # print("result Value : ", resultValue)
                if(JPresults[row][resultValue]=="Pass"):
                    sheet2.write(row+1,resultValue+2,JPresults[row][resultValue],passFormat)
                else:
                    sheet2.write(row+1,resultValue+2,JPresults[row][resultValue],failFormat)
        sheet2.autofit()
    if([] not in KOresults):
        sheet3= workbook.add_worksheet('INTL KO Test Cases')
        sheet3.write(0,0,"Test Case Description",bold)
        # sheet3.write(0,1,"Result",bold)
        sheet3.write(0,1,"Jira Links",bold)
        currentCol=2
        for values in range(len(KOresults[0])):
            sheet3.write(0,currentCol,"Dialog "+str(values+1),bold)
            currentCol+=1
        for row in range(len(KOintlExcelList)):
            sheet3.write(row+1,0,KOintlExcelList[row])
  
            sheet3.write_url(row+1,1,sampleLink.format(KOintlExcelList[row].split(':')[0].strip().replace(' ','-')),string="Open Jira Link")
            for resultValue in range(len(KOresults[row])):
                # print("result Value : ", resultValue)
                if(KOresults[row][resultValue]=="Pass"):
                    sheet3.write(row+1,resultValue+2,KOresults[row][resultValue],passFormat)
                else:
                    sheet3.write(row+1,resultValue+2,KOresults[row][resultValue],failFormat)
        sheet3.autofit()   
    if([] not in zhCNresults):
        sheet4= workbook.add_worksheet('INTL zh_CN Test Cases')
        sheet4.write(0,0,"Test Case Description",bold)
        # sheet4.write(0,1,"Result",bold)
        sheet4.write(0,1,"Jira Links",bold)
        currentCol=2
        for values in range(len(zhCNresults[0])):
            sheet4.write(0,currentCol,"Dialog "+str(values+1),bold)
            currentCol+=1
        for row in range(len(zhCNintlExcelList)):
            sheet4.write(row+1,0,zhCNintlExcelList[row])

            sheet4.write_url(row+1,1,sampleLink.format(zhCNintlExcelList[row].split(':')[0].strip().replace(' ','-')),string="Open Jira Link")
            for resultValue in range(len(zhCNresults[row])):
                # print("result Value : ", resultValue)
                if(zhCNresults[row][resultValue]=="Pass"):
                    sheet4.write(row+1,resultValue+2,zhCNresults[row][resultValue],passFormat)
                else:
                    sheet4.write(row+1,resultValue+2,zhCNresults[row][resultValue],failFormat)
        sheet4.autofit()
    if([] not in zhTWresults):
        sheet5= workbook.add_worksheet('INTL zh_TW Test Cases')
        sheet5.write(0,0,"Test Case Description",bold)
        # sheet5.write(0,1,"Result",bold)
        sheet5.write(0,1,"Jira Links",bold)
        currentCol=2
        for values in range(len(zhTWresults[0])):
            sheet4.write(0,currentCol,"Dialog "+str(values+1),bold)
            currentCol+=1
        for row in range(len(zhTWintlExcelList)):
            sheet5.write(row+1,0,zhTWintlExcelList[row])

            sheet5.write_url(row+1,1,sampleLink.format(zhTWintlExcelList[row].split(':')[0].strip().replace(' ','-')),string="Open Jira Link")
            for resultValue in range(len(zhTWresults[row])):
                # print("result Value : ", resultValue)
                if(zhTWresults[row][resultValue]=="Pass"):
                    sheet5.write(row+1,resultValue+2,zhTWresults[row][resultValue],passFormat)
                else:
                    sheet5.write(row+1,resultValue+2,zhTWresults[row][resultValue],failFormat)
        sheet5.autofit()
    workbook.close()