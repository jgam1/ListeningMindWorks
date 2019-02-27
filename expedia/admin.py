from django.contrib import admin
from expedia.models import ExpediaKeywords, ExpediaUrls

# Register your models here.

admin.site.register(ExpediaKeywords)
admin.site.register(ExpediaUrls)