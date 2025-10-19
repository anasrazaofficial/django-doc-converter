from django import forms

# Document form input by form
class DocumentForm(forms.Form):
    file = forms.FileField(required=True, label="Select a file")
