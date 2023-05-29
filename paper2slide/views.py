from .pdfExtractor import extract_data
from .summarizer import summarize_text
from .preprocessor import convert_references_section_title, extract_table, check_match, get_cleaned_text, data_reconstruction

import os
import json

from django.http import HttpResponse, HttpResponseRedirect
from django.shortcuts import render, redirect
from django.core.files.storage import FileSystemStorage

from .forms import FileUploadForm
from django.conf import settings

def index(request):
    if request.method == 'POST':
        form = FileUploadForm(request.POST, request.FILES)
        if form.is_valid():
            file = request.FILES['file']
            name, ext = os.path.splitext(file.name)
            fs = FileSystemStorage()
            pdf_file = fs.save(file.name, file)
            result = pdf_to_text(settings.MEDIA_ROOT / pdf_file)
            print(result)
            return redirect('paper2slide:choose_template', file=pdf_file) 
    else:
        form = FileUploadForm()

    return render(request, 'paper2slide/step-1.html', {'form': form})

def choose_template(request, file):
    template_list = [
        {'id': i, 'title': f'title {i}', 'thumbnail': f'https://picsum.photos/300/200?random={i}'} for i in range(10) 
    ]
    form = FileUploadForm()
    return render(request, 'paper2slide/step-2.html', {'template_list': template_list, 'form': form})

def handle_template(request):
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

    recon = data_reconstruction(processed_data)

    with open(save_path, 'w') as json_file:
        json.dump(recon, json_file)
    
    print("data saved at ", save_path)

    return recon, save_path

# TODO: Inseo
def generate_slide(paper_summary):
    """
    Generate slide with given paper summary information.

    :param paper_summary: summarized paper contents
    (The following format is just an example. You can define better format.)
    (e.g. [
        {
            'title': 'Introduction',
            'summary': ['This paper is ...', 'We discovered that ...', ...],
            'figure': ['/path/to/file', '/path/to/file', ...]
        },
        {
            'title': 'Overview of CFS and ULE',
            'subtitle': 'Linux CFS',
            'summary': ['Per-core scheduling: Linuxs CFS implements ...', ...],
            'figure': ['/path/to/file', ...]
        }
    ])
    :return: pptx slide file
    """
    pass

