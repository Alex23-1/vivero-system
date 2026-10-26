from django import forms
from .models import EnvioVivero, InventarioSobrante, Planta, VentaMenor


class PlantaForm(forms.ModelForm):

  class Meta:
    model = Planta
    fields = ['nombre', 'capacidad_bandeja', 'plantas_por_fila']
    widgets = {
        'nombre': forms.TextInput(attrs={'class': 'form-control'}),
        'capacidad_bandeja': forms.NumberInput(attrs={'class': 'form-control'}),
        'plantas_por_fila': forms.NumberInput(attrs={'class': 'form-control'}),
    }


class VentaMenorForm(forms.ModelForm):

  class Meta:
    model = VentaMenor
    fields = ['planta', 'cantidad_vendida']
    widgets = {
        'planta': forms.Select(attrs={'class': 'form-select'}),
        'cantidad_vendida': forms.NumberInput(attrs={'class': 'form-control'}),
    }


class EnvioViveroForm(forms.ModelForm):

  class Meta:
    model = EnvioVivero
    fields = ['planta', 'bandejas']
    widgets = {
        'planta': forms.Select(attrs={'class': 'form-select'}),
        'bandejas': forms.NumberInput(attrs={'class': 'form-control'}),
    }


class InventarioSobranteForm(forms.ModelForm):

  class Meta:
    model = InventarioSobrante
    fields = ['planta', 'bandejas_completas', 'filas_sueltas']
    widgets = {
        'planta': forms.Select(attrs={'class': 'form-select'}),
        'bandejas_completas': forms.NumberInput(
            attrs={'class': 'form-control', 'placeholder': 'Ej. 2'}
        ),
        'filas_sueltas': forms.NumberInput(
            attrs={'class': 'form-control', 'placeholder': 'Ej. 5'}
        ),
    }

  def __init__(self, *args, **kwargs):
    super().__init__(*args, **kwargs)
    # Si es un registro nuevo, dejamos los campos vacíos en lugar de mostrar '0'
    if not self.instance.pk:
      self.fields['bandejas_completas'].initial = ''
      self.fields['filas_sueltas'].initial = ''