from django import forms
from django.contrib.auth import get_user_model
from django.contrib.auth.forms import UserCreationForm
from .models import Visit, VisitPhoto, VisitCheckItem

User = get_user_model()


class VisitCreateForm(forms.Form):
    store_id = forms.IntegerField(widget=forms.HiddenInput)

    def __init__(self, *args, store=None, **kwargs):
        super().__init__(*args, **kwargs)
        if store:
            self.store = store


class VisitPhotoForm(forms.ModelForm):
    class Meta:
        model = VisitPhoto
        fields = ['image', 'comment']
        widgets = {
            'image': forms.FileInput(attrs={
                'accept': 'image/*',
                'class': 'form-control',
                'capture': 'environment',
            }),
            'comment': forms.TextInput(attrs={
                'class': 'form-control',
                'placeholder': 'Комментарий к фото (необязательно)',
            }),
        }

    def clean_image(self):
        image = self.cleaned_data.get('image')
        if image:
            if not image.content_type.startswith('image/'):
                raise forms.ValidationError('Можно загружать только изображения.')
        return image


class VisitNotesForm(forms.ModelForm):
    class Meta:
        model = Visit
        fields = ['notes']
        widgets = {
            'notes': forms.Textarea(attrs={
                'class': 'form-control',
                'rows': 4,
                'placeholder': 'Заметки по визиту...',
            }),
        }


class CheckItemStatusForm(forms.Form):
    item_id = forms.IntegerField(widget=forms.HiddenInput)
    status = forms.ChoiceField(choices=VisitCheckItem.STATUS_CHOICES, widget=forms.HiddenInput)


class AnalyticsFilterForm(forms.Form):
    date_from = forms.DateField(
        required=False,
        widget=forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
        label='С',
    )
    date_to = forms.DateField(
        required=False,
        widget=forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
        label='По',
    )


class UserRegistrationForm(UserCreationForm):
    first_name = forms.CharField(
        max_length=30,
        required=True,
        widget=forms.TextInput(attrs={'class': 'form-control form-control-lg', 'placeholder': 'Имя'}),
        label='Имя',
    )
    last_name = forms.CharField(
        max_length=150,
        required=True,
        widget=forms.TextInput(attrs={'class': 'form-control form-control-lg', 'placeholder': 'Фамилия'}),
        label='Фамилия',
    )
    username = forms.CharField(
        max_length=150,
        widget=forms.TextInput(attrs={'class': 'form-control form-control-lg', 'placeholder': 'Логин'}),
        label='Логин',
    )
    password1 = forms.CharField(
        widget=forms.PasswordInput(attrs={'class': 'form-control form-control-lg', 'placeholder': 'Пароль'}),
        label='Пароль',
    )
    password2 = forms.CharField(
        widget=forms.PasswordInput(attrs={'class': 'form-control form-control-lg', 'placeholder': 'Подтверждение'}),
        label='Подтвердите пароль',
    )

    class Meta:
        model = User
        fields = ('username', 'first_name', 'last_name', 'password1', 'password2')

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['username'].help_text = None
        self.fields['password1'].help_text = None
        self.fields['password2'].help_text = None

    def clean_username(self):
        username = self.cleaned_data.get('username')
        if User.objects.filter(username=username).exists():
            raise forms.ValidationError('Пользователь с таким логином уже существует.')
        return username
