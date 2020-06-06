from django.db import models

# Create your models here.
class Sale_detail(models.Model):
    customer = models.CharField(max_length=100)
    cust_c_loc = models.CharField(max_length=100)
    cust_j_loc = models.CharField(max_length=100)

class Sale1(models.Model):
    cust_code = models.CharField(max_length=100)
    cust_name = models.CharField(max_length=100)

class dpr_cust_code(models.Model):
    cust_code = models.CharField(max_length=100)
    cust_name = models.CharField(max_length=100)

class dprExcel_cellLocation(models.Model):
    longTerm_bulkerNo = models.CharField(max_length=100)
    longTerm_quantity = models.CharField(max_length=100)
    shortTerm_bulkerNo = models.CharField(max_length=100)
    shortTerm_quantity = models.CharField(max_length=100)
    brick_bulkerNo = models.CharField(max_length=100)
    brick_quantity = models.CharField(max_length=100)

class dprCumulative_cellLocation(models.Model):
    customer = models.CharField(max_length=100)
    today_quantityCell = models.CharField(max_length=100)
    cumulative_quantityCell = models.CharField(max_length=100)

class dpr_cust_code_shortTerm(models.Model):
    cust_code = models.CharField(max_length=100)
    cust_name = models.CharField(max_length=100)

class dpr_cust_codeFoc(models.Model):
    cust_code = models.CharField(max_length=100)
    cust_name = models.CharField(max_length=100)

class excel_sheet_path(models.Model):
    daily_report = models.CharField(max_length=100)
    sale_detail = models.CharField(max_length=100)
    ash_utilization = models.CharField(max_length=100)    

