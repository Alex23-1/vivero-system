from django.urls import path
from . import views

urlpatterns = [
    path('pc/', views.panel_computadora, name='panel_pc'),
    path('pc/editar/<int:pk>/', views.editar_venta, name='editar_venta'),
    path('pc/eliminar/<int:pk>/', views.eliminar_venta, name='eliminar_venta'),
    path('celular/', views.panel_celular, name='panel_celular'),
    path(
        'celular/editar/<int:pk>/',
        views.editar_sobrante,
        name='editar_sobrante',
    ),
    path(
        'celular/eliminar/<int:pk>/',
        views.eliminar_sobrante,
        name='eliminar_sobrante',
    ),
    path(
        'exportar-excel/',
        views.exportar_excel_celular,
        name='exportar_excel',
    ),
    path(
        'exportar-excel-pc/',
        views.exportar_excel_pc,
        name='exportar_excel_pc',
    ),
    path('pc/eliminar-envio/<int:pk>/', views.eliminar_envio, name='eliminar_envio'),
    path('pc/editar-envio/<int:pk>/', views.editar_envio, name='editar_envio'),
]