from django.shortcuts import render
from django.http import HttpResponse
from .models import Sale_detail
from django.contrib import messages
import pandas as pd
from datetime import date, timedelta
from babel.numbers import format_currency
import xlwings
import MySQLdb
import os.path


# ******************************************** Connect to SQL Database **************************************
def connect_sql():
    connection = MySQLdb.connect(host="localhost",user="root",passwd="",db="djnago_test")
    cur = connection.cursor()
    return cur,connection
# ***********************************    Excel sheets path **************************
# daily_report = []
# sale_detail = []
# ash_utilization = []
# cur,conn = connect_sql()
# cur.execute("SELECT * FROM dashboard_excel_sheet_path")
# for row in cur.fetchall():
#     daily_report.append(row[1])
#     sale_detail.append(row[2])
#     ash_utilization.append(row[3])
# print(daily_report)
# ******************************************** Connect to sale detail sheet **************************************
def detailsheet():
    today1 = date.today()
    month = today1.strftime("%B")
    year = today1.strftime("%Y")
    path = r'\\10.9.32.2\adm\Ash\FY 2020-21\Sale detail sheet'
    sheet = f'SALE DETAIL SHEET {month.upper()} {year}'
    full_path = os.path.join(path, sheet+'.xlsx')
    df = pd.read_excel(open(full_path,'rb'),sheet_name= month.upper() ,index_col=None, header=None,skiprows=1) 
    return df

# ******************************************** Getting DPR sale  **************************************
df = detailsheet()

# ******************************************** HOME page starts **************************************
def sale_detail(request):
    today1 = date.today()
    today = today1.strftime('%d-%m-%Y')
    yesterday1 = today1 - timedelta(days = 1)
    yesterday = yesterday1.strftime('%d-%m-%Y')
    month = today1.strftime("%B")
    year = today1.strftime("%Y")
    def adv_sale():
# ******************************************** Sale detail for current day and current month **************************************        
        sum_list = []                 # Getting current date sale in MT and no. of bulkers
        month_sum = []               # Getting current month sale in MT and no. of bulkers
        yestersum_list = []
        customers1 = []              # Getting customer code from db                            
        customers2 = []              # Getting customer name from db 
        cur,conn = connect_sql()
        cur.execute("SELECT * FROM dashboard_sale1")
        for row in cur.fetchall():
            customers1.append(int(row[1]))
            customers2.append(row[2])    
        conn.close()             
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
                sum_list.append([i, str(b),count1])
            if bifurcated_month_total !=0:
                month_sum.append([i,str(bifurcated_month_total),bifurcated_month_count])
            if sum_round !=0:  
                yestersum_list.append([i,str(sum_round),yes_count]) 

        total= df[df[6]== today][8].sum()
        month_total= df[8].sum()
        total_count= df[df[6]== today][8].count()
        month_count= df[8].count()
        yes_total= df[df[6]== yesterday][8].sum() 
        yes_total_count= df[df[6]== yesterday][8].count() 
        if total !=0:
            sum_list.append(["Total",str(round(total,2)),total_count])
        if month_total !=0:                 
            month_sum.append(["Total",str(round(month_total,2)),month_count])
        if yes_total !=0: 
            yestersum_list.append(["Total",str(round(yes_total,2)),yes_total_count])   
# ******************************************** Advance Balance **************************************
        import pythoncom
        pythoncom.CoInitialize()
        try:
            app = xlwings.App(visible=False)
            wb = app.books.open(r'\\10.9.32.2\adm\Ash\FY 2020-21\Daily Report\Daily report format.xlsx') # Daily report connection
            wb.Interactive = False
            ws = wb.sheets['advance tracking sheet']
            l1 = []                         # Getting customer name cell location from DPR- advance balance sheet
            l2= []                          # Getting advance till yesterday cell location from DPR- advance balance sheet 
            cur,conn = connect_sql()
            cur.execute("SELECT * FROM dashboard_sale_detail")
            for row in cur.fetchall():
                l1.append(row[2])
                l2.append(row[3])
            conn.close()    
            yester_bal = []                          # Getting yesterday advance balance value 
            for j in l2:
                b = ws.range(j). value
                c = b
                yester_bal.append(c)  
            df = detailsheet()

            customers = []                       # Customer code for advance balance
            cur,conn = connect_sql()
            cur.execute("SELECT * FROM dashboard_dpr_cust_code_shortTerm")
            for row in cur.fetchall():
                customers.append(int(row[1]))
            cur.execute("SELECT * FROM dashboard_dpr_cust_codeFoc")
            for row in cur.fetchall():
                customers.append(int(row[1]))        
            conn.close()

            today_sale = []
            for i in customers:
                ab = df[df[2] == i]
                amount = float((ab[ab[6]== today][19]).sum())
                today_sale.append(round(amount))
            net_bal = []   
            for (i, j) in zip(yester_bal,today_sale):
                final = round(i - j)
                net_bal.append(round(final))
                
            cust_name =[]
            for i in l1:
                b = ws.range(i). value
                cust_name.append(b)

            bal= list(zip(cust_name, net_bal))
            wb.close()
        except:
            bal = ['no']
        return (sum_list,bal,month_sum,yestersum_list)
    sum_list,bal,month_sum,yestersum_list = adv_sale()
    
    return render(request,'sale_detail.html', {'sum_list': sum_list,'bal': bal,'month_sum':month_sum,'yestersum_list':yestersum_list})
# ************************************* DPR *********************************************************************
def dpr(request):
    today1 = date.today()
    month = today1.strftime("%B")
    year = today1.strftime("%Y")
    import pythoncom
    pythoncom.CoInitialize()
    if request.method == "POST":
        userdate_date = request.POST.get('num1')
        userpath = request.POST.get('num2')
        path = r'\\10.9.32.2\adm\Ash\FY 2020-21\Sale detail sheet'
        userpath1 = f'SALE DETAIL SHEET {userpath.upper()} 2020'
        abc = os.path.join(path, userpath1+'.xlsx')
        customers1 = []
        customers2 = []
        customers3 = []
        cur,conn = connect_sql()
        cur.execute("SELECT * FROM dashboard_dpr_cust_code")
        for row in cur.fetchall():
            customers1.append(int(row[1]))
        cur.execute("SELECT * FROM dashboard_dpr_cust_code_shortTerm")
        for row in cur.fetchall():
            customers2.append(int(row[1]))
        cur.execute("SELECT * FROM dashboard_dpr_cust_codeFoc")
        for row in cur.fetchall():
            customers3.append(int(row[1]))        
        conn.close()
        df = pd.read_excel(open(abc, "rb"), sheet_name= month.upper() ,index_col=None, header=  None)
        tarik = userdate_date
    #  program for sumifs and countifs for customers1 and appending data to Dpr
        sum_list1 = []
        count_list1 =[]
        for i in customers1:
            ab = df[df[2] == i]
            a= (ab[ab[6]== tarik][8]).sum()
            b= (ab[ab[6]== tarik][8]).count()
            sum_list1.append(round(a,2))
            count_list1.append(b)

        app = xlwings.App(visible=False)
        wb = app.books.open(r'\\10.9.32.2\adm\Ash\FY 2020-21\DAILY REPORT\DAILY REPORT FORMAT.xlsx')  
        ws = wb.sheets['DPR']
        loc = []
        cur,conn = connect_sql()
        cur.execute("SELECT * FROM dashboard_dprexcel_celllocation")
        for row in cur.fetchall():
            loc.append(row)
        conn.close()
        ws.range(str(loc[0][1])).options(transpose=True).value = count_list1
        ws.range(str(loc[0][2])).options(transpose=True).value = sum_list1
            
    #  program for sumifs and countifs for customers2 and appending data to Dpr
        sum_list2 = []
        count_list2 =[]
        for i in customers2:
            ab = df[df[2] == i]
            a= (ab[ab[6]== tarik][8]).sum()
            b= (ab[ab[6]== tarik][8]).count()
            sum_list2.append(round(a,2))
            count_list2.append(b)

        ws.range(str(loc[0][3])).options(transpose=True).value = count_list2
        ws.range(str(loc[0][4])).options(transpose=True).value = sum_list2
            
    #  program for sumifs and countifs for customers3 and appending data to Dpr
        sum_list3 = []
        count_list3 =[]
        for i in customers3:
            ab = df[df[2] == i]
            a= (ab[ab[6]== tarik][8]).sum()
            b= (ab[ab[6]== tarik][8]).count()
            sum_list3.append(round(a,2))
            count_list3.append(b)

        ws.range(str(loc[0][5])).options(transpose=True).value = count_list3
        ws.range(str(loc[0][6])).options(transpose=True).value = sum_list3
        dict1=[]
        dict2=[]
        cur,conn = connect_sql()
        cur.execute("SELECT * FROM dashboard_dprCumulative_cellLocation")
        for row in cur.fetchall():
            dict1.append(str(row[2]))
            dict2.append(str(row[3]))
        conn.close()

        for i,j in zip(dict1,dict2):
            num1 = 0
            num1_new = ws.range(i).value 
            num2 = ws.range(j).value 
            ws.range(j).value = (num2+(num1_new - num1))
        
        wb.save()
        wb.close()
        messages.success(request, 'Your report has been created successfully!')     
    return render(request,'dpr.html',) 

def home(request):
    today1 = date.today()
    today = today1.strftime('%d-%m-%Y')
    yesterday1 = today1 - timedelta(days = 1)
    yesterday = yesterday1.strftime('%d-%m-%Y')
    month = today1.strftime("%B")
    year = today1.strftime("%Y")
# ************************************************* Brick_Today's sale_Yester's Sale_Month sale********************************************
    df = detailsheet()
    l = 0.00
    foc_total= df[df[9]== l][8].sum()
    today_total= round(df[df[6]== today][8].sum(),2)
    yester_total= int(df[df[6]== yesterday][8].sum())
    total1 = int(df[8].sum())
    per = round(foc_total*100/total1,2)
# ************************************************* ASH UTILIZATION ********************************************
    import pythoncom
    pythoncom.CoInitialize()
    app = xlwings.App(visible=False)
    wb = app.books.open(r'\\10.9.32.2\adm\Ash\FY 2019-20\Quantity details\MONTHWISE DETAILS 2019-20.xlsx')
    wb.Interactive = False
    ws = wb.sheets['Summary']
    a = ws.range("N21").value
    ash_utilization= round(a*100)
    wb.close()
    
    app = xlwings.App(visible=False)
    wb = app.books.open(r'\\10.9.32.2\adm\Ash\FY 2020-21\DAILY REPORT\DAILY REPORT FORMAT.xlsx') 
    wb.Interactive = False
    ws = wb.sheets['DPR']
    a= ws.range("F98").value
    pond_ash= int(a)
    wb.close()
# ******************************************** Month revenue ***************************
    df = detailsheet()
    amount = df[10].sum()
    handling = df[15].sum()
    total = amount+handling
    revenue = format_currency(int(total), 'INR',format=u'#,##0\xa0¤',currency_digits=False, locale='en_IN')

    return render(request,"home.html",{'per':per,'ash_utilization': ash_utilization,'revenue':revenue,'total1':total1,'today_total':today_total,'yester_total':yester_total,'yesterday':yesterday,'d1':today,'month':month,'pond_ash':pond_ash}) 


