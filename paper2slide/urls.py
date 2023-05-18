from django.urls import path

from . import views

app_name = 'paper2slide'
urlpatterns = [
    path('step-1', views.index, name='index'),
    path('step-2', views.handle_template, name='handle_template'),
    path('step-2/upload', views.upload_template, name='upload_template'),
    path('step-2/<str:file>', views.choose_template, name='choose_template'),
    path('step-3', views.adjust_options, name='adjust_options')
]