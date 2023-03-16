from django.urls import path

from . import views

app_name = 'paper2slide'
urlpatterns = [
    path('', views.index, name='index')
]