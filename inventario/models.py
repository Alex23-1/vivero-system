from django.db import models


class Planta(models.Model):
  nombre = models.CharField(max_length=100, unique=True)
  capacidad_bandeja = models.IntegerField(default=288)
  plantas_por_fila = models.IntegerField(default=12)

  def __str__(self):
    return self.nombre


class VentaMenor(models.Model):
  fecha = models.DateField(auto_now_add=True)
  planta = models.ForeignKey(Planta, on_delete=models.CASCADE)
  cantidad_vendida = models.IntegerField()

  def __str__(self):
    return f'{self.planta.nombre} - {self.cantidad_vendida} unidades'


class EnvioVivero(models.Model):
  fecha = models.DateField(auto_now_add=True)
  planta = models.ForeignKey(Planta, on_delete=models.CASCADE)
  bandejas = models.IntegerField()

  @property
  def total_plantas(self):
    return self.bandejas * self.planta.capacidad_bandeja

  def __str__(self):
    return f'{self.planta.nombre} - {self.bandejas} bandejas'


class InventarioSobrante(models.Model):
  fecha = models.DateField(auto_now_add=True)
  planta = models.ForeignKey(Planta, on_delete=models.CASCADE)
  bandejas_completas = models.IntegerField(default=0)
  filas_sueltas = models.IntegerField(default=0)

  @property
  def capacidad_bandeja(self):
    return self.planta.capacidad_bandeja

  @property
  def plantas_por_fila(self):
    return self.planta.plantas_por_fila

  @property
  def total_plantas(self):
    return (self.bandejas_completas * self.capacidad_bandeja) + (
        self.filas_sueltas * self.plantas_por_fila
    )

  def __str__(self):
    return (
        f'{self.planta.nombre} ({self.bandejas_completas} B +'
        f' {self.filas_sueltas} F)'
    )