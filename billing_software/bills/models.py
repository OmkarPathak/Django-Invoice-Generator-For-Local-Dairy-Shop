from django.db import models
from django import forms

class Upload(models.Model):
    excel_file    = models.FileField(('Upload File'), upload_to='uploads/%Y/%m/')
    date_uploaded = models.DateTimeField(auto_now_add=True) 

class UploadForm(forms.ModelForm):
    class Meta:
        model  = Upload
        fields = '__all__'