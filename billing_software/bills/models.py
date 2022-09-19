from django.db import models
from django import forms
from crispy_forms.helper import FormHelper

class Upload(models.Model):
    excel_file    = models.FileField(('Upload File'), upload_to='uploads/%Y/%m/')
    date_uploaded = models.DateTimeField(auto_now_add=True) 

class UploadForm(forms.ModelForm):
    class Meta:
        model  = Upload
        fields = '__all__'

class Rate(models.Model):
    id = models.AutoField(primary_key=True)
    cow_milk_rate = models.CharField("गायीचे दूध /ltr", max_length=10)
    milk_rate = models.CharField("म्हशींचे दूध /ltr", max_length=10)
    dahi_rate = models.CharField("दही /kg", max_length=10)
    ghee_rate = models.CharField("तूप /kg", max_length=10)
    updated = models.DateTimeField(auto_now_add=True) 
    
class RateForm(forms.ModelForm):
    class Meta:
        model  = Rate
        fields = '__all__'
        exclude = ['id']
    
class SingleBillForm(forms.Form):  
    cow_milk_quantity = forms.CharField(max_length=50)  
    cow_milk_rate = forms.CharField(max_length=100)
    cow_milk_cost = forms.CharField(max_length=100)
    
    milk_quantity = forms.CharField(max_length=100)
    milk_rate = forms.CharField(max_length=100)  
    milk_cost = forms.CharField(max_length=100)
    
    dahi_rate = forms.CharField(max_length=100)
    dahi_cost = forms.CharField(max_length=100)
    
    ghee_rate = forms.CharField(max_length=100)
    ghee_cost = forms.CharField(max_length=100)
    
    last_month_due = forms.CharField(max_length=100)
    total = forms.CharField(max_length=100)
    
    def __init__(self, *args, **kwargs):
        super(SingleBillForm, self).__init__(*args, **kwargs)
        self.fields['milk_quantity'].label = False
        self.fields['milk_rate'].label = False
        self.fields['milk_cost'].label = False
        self.fields['cow_milk_quantity'].label = False
        self.fields['cow_milk_rate'].label = False
        self.fields['cow_milk_cost'].label = False
        self.fields['dahi_rate'].label = False
        self.fields['dahi_cost'].label = False
        self.fields['ghee_rate'].label = False
        self.fields['ghee_cost'].label = False
        self.fields['last_month_due'].label = False
        self.fields['total'].label = False