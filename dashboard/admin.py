from django.contrib import admin
from .models import Sale_detail
from .models import Sale1
from .models import dpr_cust_code
from .models import dpr_cust_code_shortTerm
from .models import dpr_cust_codeFoc
from .models import dprExcel_cellLocation
from .models import dprCumulative_cellLocation
from .models import excel_sheet_path

# Register your models here.
admin.site.register(Sale_detail)
admin.site.register(Sale1)
admin.site.register(dpr_cust_code)
admin.site.register(dpr_cust_code_shortTerm)
admin.site.register(dpr_cust_codeFoc)
admin.site.register(dprExcel_cellLocation)
admin.site.register(dprCumulative_cellLocation)
admin.site.register(excel_sheet_path)