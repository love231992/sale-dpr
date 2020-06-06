import pandas as pd
from datetime import date, timedelta
import time
import xlwings
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
import numpy as np

def data_gsheet():
    # today1 = date.today()
    # today = today1.strftime('%d-%m-%Y')
    # yesterday1 = today1 - timedelta(days = 1)
    # yesterday = yesterday1.strftime('%d-%m-%Y')
    # month = today1.strftime("%B")
    # year = today1.strftime("%Y")
    today = "15-04-2020"
    yesterday = "21-03-2020"

    def detailsheet():
        today1 = date.today()
        month = today1.strftime("%B")
        # year = today1.strftime("%Y")
        # path = r'\\10.9.32.2\adm\Ash\FY 2019-20\Sale detail sheet'
        # sheet = f'SALE DETAIL SHEET {month.upper()} {year}'
        # full_path = os.path.join(path, sheet+'.xlsx')
        df = pd.read_excel(open(r"\\10.9.32.2\adm\Ash\FY 2020-21\Sale Detail Sheet\SALE DETAIL SHEET APRIL 2020.xlsx",'rb'),sheet_name= month.upper() ,index_col=None, header=None,skiprows=1) 
        return df

    def adv_sale():
    # ******************************************** Sale detail for current day and current month **************************************        
        sum_list= []                 # Getting current date sale in MT and no. of bulkers
        month_sum = []               # Getting current month sale in MT and no. of bulkers
        yestersum_list = []
        # customers1 = []              # Getting customer code from db                            
        # customers2 = []              # Getting customer name from db 
        customers1 = [20,31,28,27,17,46,18,13,15,14,37,100125,100051,100062,100087,100072,100071,100070,100057,100056,100066,100068,100086,100091,100103,100126,100131,100145,100150,100152,100140,100165,100180]
        customers2 = [ 'Asian Cements','Asian Fine', 'UTCL Roorkee', 'UTCL Bagheri','UTCL Panipat','UTCL Sikandrabad','UTCL Bathinda','Ambuja Nalagargh','Ambuja Ropar','Ambuja Roorkee','Ambuja Dadri','ACC LTD','Fateh','Everest','Hemkund Sahib','Rakesh kumar','Amritsaria','Jai Shiv shankar','Ramjee','Paras','Manju','Sachin','R.S.Green','BTS','S.A.Bricks','Royal','M.M. Concrete','Fairmont','Aniket','A One','ONS','NTC','Guru Teg Bhadar']
        # cur,conn = connect_sql()
        # cur.execute("SELECT * FROM dashboard_sale1")
        # for row in cur.fetchall():
        #     customers1.append(int(row[1]))
        #     customers2.append(row[2])    
        # conn.close()             
        customers2_length = range(0,len(customers2))
        df = detailsheet()
        for (i,j) in zip(customers1, customers2_length):
            ab = df[df[2] == i]
            bifurcated_month_total = round(ab[8].sum(),2) 
            bifurcated_month_count = ab[8].count() 
            a= (ab[ab[6]== today][8]).sum()
            count1= (ab[ab[6]== today][8]).count()
            b = round(a,2)
            yes_sum = (ab[ab[6]== yesterday][8]).sum() 
            yes_count = (ab[ab[6]== yesterday][8]).count() 
            sum_round = round(yes_sum,2)
            i = customers2[j]
            if b != 0:
                sum_list.append([i, str(b),str(count1)])
            if bifurcated_month_total !=0:
                month_sum.append([i,str(bifurcated_month_total),str(bifurcated_month_count)])
            if sum_round !=0:  
                yestersum_list.append([i,str(sum_round),str(yes_count)]) 

        total= df[df[6]== today][8].sum()
        month_total= df[8].sum()
        total_count= df[df[6]== today][8].count()
        month_count= df[8].count()
        yes_total= df[df[6]== yesterday][8].sum() 
        yes_total_count= df[df[6]== yesterday][8].count() 
        if total !=0:
            sum_list.append(["Total",str(round(total,2)),str(total_count)])
        if month_total !=0:                 
            month_sum.append(["Total",str(round(month_total,2)),str(month_count)])
        if yes_total !=0: 
            yestersum_list.append(["Total",str(round(yes_total,2)),str(yes_total_count)]) 
        return (sum_list,month_sum,yestersum_list)    

    current,month1,yester1 = adv_sale()
    # print(type(current))

    scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

    creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)

    client = gspread.authorize(creds)

    # sheet1 = client.open("Tutorial").Sheet2
    spreadsheet = client.open("Tutorial")
    sheet1 = spreadsheet.worksheet("Today")
    sheet2 = spreadsheet.worksheet("Demo")
    # sheet2 = spreadsheet.get_worksheet(1)
    # sheet3 = spreadsheet.get_worksheet(2)

    spreadsheet.del_worksheet(sheet1)
    sheet1 = spreadsheet.add_worksheet(title="Today",rows="100",cols="50")
    spreadsheet.del_worksheet(sheet2)
    sheet2 = spreadsheet.add_worksheet(title="Demo",rows="100",cols="50")
    # spreadsheet.del_worksheet(sheet3)
    # sheet3 = spreadsheet.add_worksheet(title="Yesterday",rows="100",cols="50")

    for pos,i in enumerate(current):
        sheet1.insert_row(current[pos], pos+1)

    # time.sleep(5)
    # for pos,i in enumerate(month1):
    #     sheet2.insert_row(month1[pos], pos+1)
    # time.sleep(5)
    # for pos,i in enumerate(yester1):
    #     sheet3.insert_row(yester1[pos], pos+1)
        

    # print(sum1[0][0])
    # spreadsheet.del_worksheet(sheet)
while True:
    data_gsheet()
    time.sleep(60)