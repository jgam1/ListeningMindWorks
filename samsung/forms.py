from django import forms
from .models import UploadFileModel

class UploadFileForm(forms.ModelForm):
    class Meta:  # 이 폼을 만들기 위해 어떤 model이 쓰여야 하는지 장고에게 알려주는 구문 
        model = UploadFileModel
        fields = ('upload_file',)
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['upload_file'].required = False #file 값이 없더라도 view의 유효성 검사에서 오류를 발생시키지 않도록 함

class FileFieldForm(forms.Form):
    file_field = forms.FileField()