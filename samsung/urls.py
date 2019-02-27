from django.urls import path
from django.conf import settings
from django.conf.urls.static import static
from . import views

app_name = 'samsung'
urlpatterns = [
    #/samsung/
    path('', views.upload_file, name='samsung_index'),
    path('excel_export', views.excel_export, name = 'excel_export'),

] + static(settings.MEDIA_URL, document_root = settings.MEDIA_ROOT)
# for file upload, it may have the static function