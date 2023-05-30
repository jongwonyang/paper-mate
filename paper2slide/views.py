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
    """
    Divide given paper (pdf file) into sections,
    (e.g. Inroduction, Background, Evaluation ...)
    and make it to text format.

    :param pdf_file: input paper pdf file
    :return: proper format for describing papaer
    (The following format is just an example. You can define better format.)
    (e.g. [
        {
            'title': 'Introduction',
            'body': 'Operating system kernel schedulers are ...'
        },
        {
            'title': 'Overview of CFS and ULE',
            'body': [
                {
                    'title': 'Linux CFS',
                    'body': 'Per-core scheduling: Linuxs CFS implements ...'
                },
                {
                    'title': 'FreeBSD ULE',
                    'body': 'Per-core scheduling: ULE uses two runqueues ...'
                }
            ]
        },
        ...
    ])
    """
    extracted_data = extract_data(pdf_file)
    paragraphs = extracted_data["paragraphs"]
    tables = extracted_data["tables"]

    output = []
    reference_flag = 0
    for content in paragraphs:
        if reference_flag == 0:
            if convert_references_section_title(content["content"]) == "REFERENCES":
                output.append({"role":"sectionHeading",
                    "content": "REFERENCES"})
                reference_flag = 1
            elif content["role"] == "sectionHeading" or content["role"] == None:
                if content["content"].upper().strip() == "ACKNOWLEDGEMENTS":
                    break;
                elif content["content"].upper().strip() == "ABSTRACT":
                    output.append({"role":"sectionHeading",
                            "content": "ABSTRACT"})
                elif len(output)>0 and content["role"] == None and output[-1]["role"]==None:
                    if len(content["content"]) > 10:
                        output[-1]["content"] = output[-1]["content"]+" "+content["content"]
                else:
                    output.append({"role":content["role"],
                            "content":content["content"]})
        elif reference_flag == 1:
            if output[-1]["content"] == "REFERENCES":
                output.append({"role":content["role"],"content":content["content"]})
            elif content["role"] == "sectionHeading":
                reference_flag == 0
            else:
                output[-1]["content"] = output[-1]["content"]+" "+content["content"]

    for i in range(len(output)):
        output[i]["content"] = get_cleaned_text(output[i]["content"])
        # et al. 때문에 .으로 문장을 구분하는 방식에 어려움 존재 -> 먼저 제거 후 다시 삽입

    output = summarize_text(output)

    processed_data = {}
    processed_data["sentences"] = output
    processed_data["tables"] = extract_table(tables)

    for i in range(len(processed_data["sentences"])):
        processed_data["sentences"][i]["tables"] = check_match(processed_data["sentences"][i]["content"], "table")
    for i in range(len(processed_data["sentences"])):
        processed_data["sentences"][i]["figures"] = check_match(processed_data["sentences"][i]["content"], 'figure')

    processed_data["sentences"] = data_reconstruction(processed_data)
    
    all_text = ""
    for content in processed_data["sentences"]:
        if content["content"] is not None:
            all_text = all_text + " " + content["content"]
    keywords = extract_keywords_from_paragraph(all_text)

    processed_data["keywords"] = keywords

    with open(save_path, 'w') as json_file:
        json.dump(processed_data, json_file)

    print("data saved at ", save_path)

    return processed_data, save_path

# TODO: Inseo


def generate_slide(paper_summary):

