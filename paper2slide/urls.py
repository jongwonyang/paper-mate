from django.urls import path

from . import views

app_name = 'paper2slide'
urlpatterns = [
    path('step-1', views.index, name='index')
]