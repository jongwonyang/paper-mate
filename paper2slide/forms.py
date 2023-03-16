from django import forms

class PaperUploadForm(forms.Form):
    file = forms.FileField()