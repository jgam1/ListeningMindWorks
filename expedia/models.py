from django.db import models

# Create your models here.
class ExpediaKeywords(models.Model):
    keyword = models.CharField(max_length=100)
    def __str__(self):
        return self.keyword
    
class ExpediaUrls(models.Model):
    url = models.CharField(max_length=100)
    
    def __str__(self):
        return self.url

