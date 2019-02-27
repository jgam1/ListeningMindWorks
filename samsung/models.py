from django.db import models

# Create your models here.

class UploadFileModel(models.Model):
    templates = 'samsung_index.html'
    upload_file = models.FileField(null=True)