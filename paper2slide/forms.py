from django import forms

class FileUploadForm(forms.Form):
    file = forms.FileField()

FONT_CHOICES = (
    ('Arial', 'Arial'),
    ('Bahnschrift', 'Bahnschrift'),
    ('Calibri', 'Calibri'),
    ('Courier New', 'Courier New'),
    ('Euphemia', 'Euphemia'),
)
RATIO_CHOICES = (
    (False, '4:3'),
    (True, '16:9')
)
class SlideOptionForm(forms.Form):
    title = forms.CharField(label='제목', max_length=100)
    username = forms.CharField(label='사용자명', max_length=100)
    title_font = forms.ChoiceField(label='제목 폰트', choices=FONT_CHOICES)
    subtitle_font = forms.ChoiceField(label='부제목 폰트', choices=FONT_CHOICES)
    body_font = forms.ChoiceField(label='본문 폰트', choices=FONT_CHOICES)
    spacing = forms.IntegerField(label='자간', max_value=100)
    ratio = forms.ChoiceField(label='화면 비율', choices=RATIO_CHOICES)