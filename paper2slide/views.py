from .pdfExtractor import extract_data
from .summarizer import summarize_text, extract_keywords_from_paragraph
from .preprocessor import convert_references_section_title, extract_table, check_match, get_cleaned_text, data_reconstruction

import os
import json

from django.http import HttpResponse, HttpResponseRedirect
from django.shortcuts import render, redirect
from django.core.files.storage import FileSystemStorage

from .forms import FileUploadForm
from django.conf import settings

import collections
import collections.abc
import json
import collections
import collections.abc
import json
import win32com.client
import pandas as pd

# TODO: Jongwon


def index(request):
    if request.method == 'POST':
        form = FileUploadForm(request.POST, request.FILES)
        if form.is_valid():
            file = request.FILES['file']
            fs = FileSystemStorage()
            pdf_file = fs.save(file.name, file)
            return redirect('paper2slide:process_pdf', pdf_file_name=pdf_file) 
    else:
        form = FileUploadForm()

    return render(request, 'paper2slide/step-1.html', {'form': form})

def process_pdf(request, pdf_file_name):
    return render(request, 'paper2slide/process-pdf.html', {'file': pdf_file_name}) 

def handle_template(request, summary_json_file):
    if request.method == 'POST':
        # create slide here
        with open(settings.MEDIA_ROOT / summary_json_file, 'r') as file:
            summary_text = file.read()
            generate_slide(summary_text)
        return redirect('paper2slide:adjust_options')
    template_list = [
        {'id': i, 'title': f'title {i}', 'thumbnail': f'https://picsum.photos/300/200?random={i}'} for i in range(10) 
    ]
    form = FileUploadForm()
    return render(request, 'paper2slide/step-2.html', {'template_list': template_list, 'form': form})

# TODO: remove
def apply_template(request):
    # Create slide here
    return redirect('paper2slide:adjust_options')

def upload_template(request):
    return HttpResponse("Upload template!")

def adjust_options(request):
    filename = "sample.pps"
    return render(request, 'paper2slide/step-3.html', {'filename': filename})

# TODO: Heejae
def pdf_to_text(pdf_file, save_path):
    