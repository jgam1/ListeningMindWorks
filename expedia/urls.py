from django.urls import path
from . import views

app_name = 'expedia'
urlpatterns = [
    #/expedia/
    path('', views.ExpediaModelView.as_view(), name='index'),

    path('expediakeywords_list/', views.KeywordsList.as_view(), name='expediakeywords_list'),

    path('expediaurls/', views.UrlsList.as_view(), name='expediaurls_list'),

    path('process/', views.process, name='process'),

    path('results/', views.results.as_view(), name='results')

]