from django.urls import path

from . import views

app_name = 'paper2slide'
urlpatterns = [
    path('step-1', views.index, name='index'),
    path('process-pdf/<str:pdf_file_name>', views.process_pdf, name='process_pdf'),
    path('step-2/<str:summary_json_file>', views.handle_template, name='handle_template'),
    path('step-2/upload/<str:summary_json_file>', views.upload_template, name='upload_template'),
    path('step-3/<str:pptx_file_name>', views.adjust_options, name='adjust_options'),
    path('step-4',views.download_pptx, name='download_pptx'),
]