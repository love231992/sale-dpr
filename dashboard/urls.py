from django.urls import path
from . import views


urlpatterns = [
    # path('', views.index, name='index'),
    path('dpr/dpr',views.dpr, name='dpr'),
    path('home/',views.home, name='home'),
    # path('report',views.report, name='report'),
    path('sale_detail/',views.sale_detail, name='sale_detail')
]