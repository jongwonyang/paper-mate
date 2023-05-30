from .pdfExtractor import extract_data
from .summarizer import summarize_text, extract_keywords_from_paragraph
from .preprocessor import convert_references_section_title, extract_table, find_pattern_match_position, get_cleaned_text, data_reconstruction

import os
import json

from django.http import HttpResponse, HttpResponseRedirect
from django.shortcuts import render, redirect
from django.core.files.storage import FileSystemStorage

from .forms import FileUploadForm, SlideOptionForm
from django.conf import settings

import collections
import collections.abc
import json
import collections
import collections.abc
import json
import win32com.client
import pandas as pd

import win32com.client
import pythoncom

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

def adjust_options(request, pptx_file_name):
    name, _ = os.path.splitext(pptx_file_name)
    pythoncom.CoInitialize()
    powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
    powerpoint.Visible = True 
    deck = powerpoint.Presentations.Open(settings.MEDIA_ROOT / pptx_file_name)
    deck.SaveAs(settings.MEDIA_ROOT / f'{name}.pdf', 32)
    deck.Close()
    powerpoint.Quit()
    form = SlideOptionForm()
    return render(request, 'paper2slide/step-3.html', {'pdf_file_name': f'{name}.pdf', 'form': form})

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


    processed_data["sentences"] = data_reconstruction(processed_data)
    
    for i in range(len(processed_data["sentences"])):
        if 'summarized' in processed_data["sentences"][i]:
            if processed_data["sentences"][i]["content"] is not None:
                processed_data["sentences"][i]["tables"] = find_pattern_match_position(processed_data["sentences"][i]["content"],processed_data["sentences"][i]["summarized"], "table")
            else: 
                processed_data["sentences"][i]["tables"] = None
        else:
            processed_data["sentences"][i]["tables"] = None
    for i in range(len(processed_data["sentences"])):
        if 'summarized' in processed_data["sentences"][i]:
            if processed_data["sentences"][i]["content"] is not None:
                processed_data["sentences"][i]["figures"] = find_pattern_match_position(processed_data["sentences"][i]["content"],processed_data["sentences"][i]["summarized"], 'figure')
            else: 
                processed_data["sentences"][i]["figures"] = None
        else:
            processed_data["sentences"][i]["figures"] = None


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


slide_layout = {"Title": 0,
                "Title and Content": 1,
                "Section Header": 2,
                "Two Content": 3,
                "Comparison": 4,
                "Title Only": 5,
                "Blank": 6,
                "Contnet with Caption": 7,
                "Picture with Caption": 8}


def generate_slide(paper_summary):

    def RowCol(csv_file_address):
        results = pd.read_csv(csv_file_address)
        return len(results.axes[0]), len(results.axes[1])


    def dataFrame_csv(csv_file_address):
        return pd.read_csv(csv_file_address)
    
    slide_layout = {"Title": 1,
                "Title and Content": 2,
                "Section Header": 3,
                "Two Content": 4,
                "Comparison": 5,
                "Title Only": 6,
                "Blank": 7,
                "Contnet with Caption": 8,
                "Picture with Caption": 9}

    def generate_slide(paper_summary):

        paper_summary = json.loads(paper_summary)
        name = paper_summary["UserInfo"]["UserID"] + \
            "_" + paper_summary["UserInfo"]["RequestID"]

        PPTApp = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
        presentation = PPTApp.Presentations.Add()

        if(paper_summary["UserInfo"]["template"] != "basic"):
            template_path = paper_summary["UserInfo"]["template"]
            presentation.ApplyTemplate(template_path)

        for slide_info in paper_summary["slideInfos"]:
            layout = presentation.Designs.Item(1).SlideMaster.CustomLayouts.Item(slide_layout[slide_info["layout"]])
            new_slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, layout)

            for index1, placeholder_info in enumerate(slide_info["placeholder"]):
                
                placeholder_type = placeholder_info["type"]
                if (placeholder_type == "title"):
                    new_slide.Shapes.Item(index1+1).TextFrame.TextRange.Text = placeholder_info["content"][0]["item"]

                elif (placeholder_type == "subtitle"):
                    new_slide.Shapes.Item(index1+1).TextFrame.TextRange.Text = placeholder_info["content"][0]["item"]
                
                elif (placeholder_type == 'contents placeholder'):
                    body = new_slide.Shapes.Item(index1+1).TextFrame.TextRange
                    Textinsertion = 0
                    for index2, content in enumerate(placeholder_info["content"]):
                        if (content["type"] == "text"):
                            this_paragraph = body.Paragraphs(index2+1)
                            if (placeholder_info["numberOfTextcontent"] == Textinsertion +1):
                                this_paragraph.Text = content["item"]
                                Textinsertion = Textinsertion +1
                            else:
                                this_paragraph.Text = content["item"] + "\n"
                                Textinsertion = Textinsertion +1
                            #this_paragraph.IndentLevel = content["level"] +1
                            this_paragraph.ParagraphFormat.Bullet.Visible = content["bullet"]


                        elif (content["type"] == "table"):
                            table_csv_file = placeholder_info["content"][index2]["item"]
                            csv_df = pandas.read_csv(table_csv_file)
                            
                            header = pandas.DataFrame(csv_df.columns).transpose()
                            numRow, numCol = csv_processor.RowCol(table_csv_file)
                            numRow = numRow + 1
                            
                            table_shape = new_slide.Shapes.AddTable(NumRows = numRow, NumColumns = numCol,
                                                                    Left = placeholder_info["content"][index2]["Left"],
                                                                    Top = placeholder_info["content"][index2]["Top"],
                                                                    Width = placeholder_info["content"][index2]["Width"],
                                                                    Height = placeholder_info["content"][index2]["Height"])
                            table = table_shape.Table
                            this_shape = new_slide.Shapes.Item(index1+1)
                            this_width = this_shape.Width
                            this_height = this_shape.Height

                            print(f"height: {this_height}")

                            if this_height > placeholder_info["content"][index2]["Height"]:
                                divider = 2
                                while (this_height > placeholder_info["content"][index2]["Height"]):

                                    table = table_shape.Table
                                    this_shape = new_slide.Shapes.Item(index1+1)
                                    this_width = this_shape.Width
                                    this_height = this_shape.Height

                                    print("===================")
                                    print(f"divider: {divider}")
                                    
                                    print(f"this_shape: {this_shape}")
                                    
                                    if this_shape.HasTable:
                                        this_shape.Delete()
                                        print("this_shape is deleted")
                                    

                                    df_list = list()
                                    whole_df = csv_df.copy()
                                    pivot = len(whole_df.index)//divider
                                    remainder = len(whole_df.index)%divider
                                    header.columns = range(header.shape[1])
                                    if divider >= 2:
                                        for i in range(divider):
                                                if(i == 0):
                                                    tmp = whole_df.copy().truncate(after = pivot)
                                                    tmp.columns = range(tmp.shape[1])
                                                    df_list.append(pandas.concat([header, tmp], axis=0))
                                                elif(i * pivot < (i+1)* pivot -1):
                                                    tmp = whole_df.copy().truncate(before = i * pivot, after = (i+1)* pivot-1)
                                                    tmp.columns = range(tmp.shape[1])
                                                    df_list.append(pandas.concat([header, tmp]))
                                        if remainder > 0 :
                                            tmp = whole_df.copy().truncate(before = divider * pivot, after = len(whole_df.index))
                                            tmp.columns = range(tmp.shape[1])
                                            df_list.append(pandas.concat([header, tmp]))
                                    else:
                                        if remainder == 0:
                                            tmp = whole_df.copy().truncate(after = pivot -1)
                                            tmp.columns = range(tmp.shape[1])
                                            df_list.append(pandas.concat([header, tmp]))
                                            tmp = header,whole_df.copy().truncate(before = pivot)
                                            tmp.columns = range(tmp.shape[1])
                                            df_list.append(pandas.concat([header, tmp]))
                                        else:
                                            tmp = whole_df.copy().truncate(after = pivot)
                                            tmp.columns = range(tmp.shape[1])
                                            df_list.append(pandas.concat([header, tmp]))
                                            tmp = header,whole_df.copy().truncate(before = pivot+1)
                                            tmp.columns = range(tmp.shape[1])
                                            df_list.append(pandas.concat([header, tmp]))


                                    first_df = df_list[0]

                                    print(f"first_df")
                                    print(first_df)

                                    table_shape = new_slide.Shapes.AddTable(NumRows = len(first_df.index), NumColumns = len(first_df.columns)
                                                                            # ,
                                                                            # Left = placeholder_info["content"][0]["Left"],
                                                                            # Top = placeholder_info["content"][0]["Top"],
                                                                            # Width = placeholder_info["content"][0]["Width"],
                                                                            # Height = placeholder_info["content"][0]["Height"]
                                                                            )
                                    table = table_shape.Table
                                    # print(f"firstRow: {first_df.iloc[0]}")
                                    for k in range(len(first_df.index)):
                                        for l in range(len(first_df.columns)):
                                            # cell = table.Cell(k+1, l+1)
                                            table.Cell(k+1, l+1).Shape.TextFrame.TextRange.Text = str(first_df.iloc[k,l])

                                    this_height = new_slide.Shapes.Item(index1+1).Height
                                    print(f"height: {this_height}")
                                    if this_height <= placeholder_info["content"][index2]["Height"]:
                                        for j in range(1, divider):
                                            new_slide.Copy()
                                            presentation.Slides.Paste(presentation.Slides.Count + 1)
                                            new_slide = presentation.Slides(presentation.Slides.Count)
                                            for element in range(new_slide.Shapes.Count):
                                                t_shape = new_slide.Shapes.Item(element+1)
                                                if t_shape.HasTable:
                                                    t_shape.Delete()
                                                    this_df = df_list[j]
                                                    table_shape = new_slide.Shapes.AddTable(NumRows = len(this_df.index), NumColumns = len(this_df.columns)
                                                                                            # ,
                                                                                            # Left = placeholder_info["content"][0]["Left"],
                                                                                            # Top = placeholder_info["content"][0]["Top"],
                                                                                            # Width = placeholder_info["content"][0]["Width"],
                                                                                            # Height = placeholder_info["content"][0]["Height"]
                                                                                            )
                                                    table = table_shape.Table
                                                    for m in range(len(this_df.index)):
                                                        for n in range(len(this_df.columns)):
                                                            cell = table.Cell(m+1, n+1)
                                                            cell.Shape.TextFrame.TextRange.Text = str(this_df.iloc[m,n])

                                    else: divider = divider +1
                            else:
                                header.columns = range(header.shape[1])
                                tmp = csv_df
                                tmp.columns = range(tmp.shape[1])
                                csv_df = pandas.concat([header, tmp])
                                for k in range(len(first_df.index)):
                                    for l in range(len(first_df.columns)):
                                        # cell = table.Cell(k+1, l+1)
                                        table.Cell(k+1, l+1).Shape.TextFrame.TextRange.Text = str(csv_df.iloc[k,l])

                        elif (content["type"] == "image"):
                            new_slide.Shapes.AddPicture(FileName = placeholder_info["content"][index2]["item"],
                                                                    LinkToFile = False,
                                                                    SaveWithDocument = True,
                                                                    Left = placeholder_info["content"][index2]["Left"],
                                                                    Top = placeholder_info["content"][index2]["Top"],
                                                                    Width = placeholder_info["content"][index2]["Width"],
                                                                    Height = placeholder_info["content"][index2]["Height"])
                                
                
                elif (placeholder_type == "text placeholder"):
                    body = new_slide.Shapes.Item(index1+1).TextFrame.TextRange
                    for index2, content in enumerate(placeholder_info["content"]):
                        this_paragraph = body.Paragraphs(index2+1)
                        if (placeholder_info["numberOfTextcontent"] == index2+1):
                                this_paragraph.Text = content["item"]
                        else:
                                this_paragraph.Text = content["item"] + "\n"

                elif (placeholder_type == "picture placeholder"):
                    new_slide.Shapes.AddPicture(FileName = placeholder_info["content"][0]["item"],
                                                    LinkToFile = False,
                                                    SaveWithDocument = True,
                                                    Left = placeholder_info["content"][0]["Left"],
                                                    Top = placeholder_info["content"][0]["Top"],
                                                    Width = placeholder_info["content"][0]["Width"],
                                                    Height = placeholder_info["content"][0]["Height"])

                elif (placeholder_type == "table placeholder"):
                    table_csv_file = placeholder_info["content"][0]["item"]
                    csv_df = pandas.read_csv(table_csv_file)
                    
                    header = pandas.DataFrame(csv_df.columns).transpose()

                    numRow, numCol = csv_processor.RowCol(table_csv_file)
                    numRow = numRow + 1
                    table_shape = new_slide.Shapes.AddTable(NumRows = numRow, NumColumns = numCol,
                                                            Left = placeholder_info["content"][0]["Left"],
                                                            Top = placeholder_info["content"][0]["Top"],
                                                            Width = placeholder_info["content"][0]["Width"],
                                                            Height = placeholder_info["content"][0]["Height"])
                    table = table_shape.Table
                    this_shape = new_slide.Shapes.Item(index1+1)
                    this_width = this_shape.Width
                    this_height = this_shape.Height


                    if this_height > placeholder_info["content"][0]["Height"]:
                        divider = 2

                        print(f"this_shape: {this_shape}")

                        this_shape.Delete

                        print("this_shape is deleted")

                        while (this_height > placeholder_info["content"][0]["Height"]):
                            df_list = list()
                            whole_df = csv_df.copy()
                            pivot = len(whole_df.index)//divider
                            remainder = len(whole_df.index)%divider
                            header.columns = range(header.shape[1])
                            if divider > 2:
                                for i in range(divider):
                                    if(i * pivot < (i+1)* pivot -1): 
                                        tmp = whole_df.copy().truncate(before = i * pivot, after = (i+1)* pivot-1)
                                        tmp.columns = range(tmp.shape[1])
                                        df_list.append(pandas.concat([header, ]))
                                if remainder > 0 :
                                    tmp = whole_df.copy().truncate(before = divider * pivot, after = len(whole_df.index))
                                    tmp.columns = range(tmp.shape[1])
                                    df_list.append(pandas.concat([header, tmp]))
                            else:
                                if remainder == 0:
                                    tmp = whole_df.copy().truncate(after = pivot -1)
                                    tmp.columns = range(tmp.shape[1])
                                    df_list.append(pandas.concat([header, tmp]))
                                    tmp = whole_df.copy().truncate(before = pivot)
                                    tmp.columns = range(tmp.shape[1])
                                    df_list.append(pandas.concat([header, tmp]))
                                else:
                                    tmp = whole_df.copy().truncate(after = pivot)
                                    tmp.columns = range(tmp.shape[1])
                                    df_list.append(pandas.concat([header,tmp]))
                                    tmp = whole_df.copy().truncate(before = pivot+1)
                                    tmp.columns = range(tmp.shape[1])
                                    df_list.append(pandas.concat([header,tmp]))

                            first_df = df_list[0]
                            table_shape = new_slide.Shapes.AddTable(NumRows = len(first_df.index), NumColumns = len(first_df.columns),
                                                                    Left = placeholder_info["content"][0]["Left"],
                                                                    Top = placeholder_info["content"][0]["Top"],
                                                                    Width = placeholder_info["content"][0]["Width"],
                                                                    Height = placeholder_info["content"][0]["Height"])
                            for k in range(len(first_df.index)):
                                for l in range(len(first_df.columns)):
                                    cell = table.Cell(k+1, l+1)
                                    cell.Shape.TextFrame.TextRange.Text = first_df.iloc[k,l]

                            this_height = new_slide.Shapes.Item(index1+1).Height
                            if this_height <= placeholder_info["content"][0]["Height"]:
                                for j in range(1, divider):
                                    new_slide.Copy()
                                    presentation.Slides.Paste(presentation.Slides.Count + 1)
                                    new_slide = presentation.Slides(presentation.Slides.Count)
                                    for element in range(new_slide.Shapes.Count):
                                        t_shape = new_slide.Shapes.Item(element+1)
                                        if t_shape.HasTable:
                                            t_shape.Delete
                                            this_df = df_list[j]
                                            table_shape = new_slide.Shapes.AddTable(NumRows = len(this_df.index), NumColumns = len(this_df.columns),
                                                                    Left = placeholder_info["content"][0]["Left"],
                                                                    Top = placeholder_info["content"][0]["Top"],
                                                                    Width = placeholder_info["content"][0]["Width"],
                                                                    Height = placeholder_info["content"][0]["Height"])
                                            for k in range(len(this_df.index)):
                                                for l in range(len(this_df.columns)):
                                                    cell = table.Cell(k+1, l+1)
                                                    cell.Shape.TextFrame.TextRange.Text = this_df.iloc[k,l]

                            else: divider = divider +1

                    else:
                        header.columns = range(header.shape[1])
                        csv_df.columns = range(csv_df.shape[1])
                        csv_df = pandas.concat([header, csv_df])
                        for k in range(len(first_df.index)):
                            for l in range(len(first_df.columns)):
                                # cell = table.Cell(k+1, l+1)
                                table.Cell(k+1, l+1).Shape.TextFrame.TextRange.Text = str(csv_df.iloc[k,l])

                else: #None
                    Textinsertion = 0
                    for index2, content in enumerate(placeholder_info["content"]):
                        shape = new_slide.Shapes.AddShape( 1 ,Left = placeholder_info["content"][0]["Left"],
                                                    Top = placeholder_info["content"][0]["Top"],
                                                    Width = placeholder_info["content"][0]["Width"],
                                                    Height = placeholder_info["content"][0]["Height"])
                        shape.Fill.Transparency = 1.0
                        shape.Line.Transparency = 1.0
                        if (content["type"] == "text"):
                            body = new_slide.Shapes.Item(index1+1+index2).TextFrame.TextRange
                            this_paragraph = body.Paragraphs(index2+1)
                            
                            this_paragraph.Text = content["item"]
                                
                            
                            #this_paragraph.IndentLevel = content["level"] +1
                            this_paragraph.ParagraphFormat.Bullet.Visible = content["bullet"]

                        elif (content["type"] == "table"):
                            table_csv_file = placeholder_info["content"][0]["item"]
                            csv_df = pandas.read_csv(table_csv_file)

                            header = pandas.DataFrame(csv_df.columns).transpose()

                            numRow, numCol = csv_processor.RowCol(table_csv_file)
                            numRow = numRow + 1
                            table_shape = new_slide.Shapes.AddTable(NumRows = numRow, NumColumns = numCol,
                                                                    Left = placeholder_info["content"][0]["Left"],
                                                                    Top = placeholder_info["content"][0]["Top"],
                                                                    Width = placeholder_info["content"][0]["Width"],
                                                                    Height = placeholder_info["content"][0]["Height"])
                            table = table_shape.Table
                            this_shape = new_slide.Shapes.Item(index1+1)
                            this_width = this_shape.Width
                            this_height = this_shape.Height

                            print(f"height: {this_height}")

                            if this_height > placeholder_info["content"][0]["Height"]:
                                divider = 2
                                while (this_height > placeholder_info["content"][0]["Height"]):

                                    table = table_shape.Table
                                    this_shape = new_slide.Shapes.Item(index1+1)
                                    this_width = this_shape.Width
                                    this_height = this_shape.Height

                                    print("===================")
                                    print(f"divider: {divider}")
                                    
                                    print(f"this_shape: {this_shape}")
                                    
                                    if this_shape.HasTable:
                                        this_shape.Delete()
                                        print("this_shape is deleted")
                                    
                                    df_list = list()
                                    whole_df = csv_df.copy()
                                    pivot = len(whole_df.index)//divider
                                    remainder = len(whole_df.index)%divider
                                    
                                    header.columns = range(header.shape[1])
                                    if divider > 2:
                                        for i in range(divider):
                                                if(i * pivot < (i+1)* pivot -1):
                                                    tmp = whole_df.copy().truncate(before = i * pivot, after = (i+1)* pivot-1)
                                                    tmp.columns = range(tmp.shape[1])
                                                    df_list.append(pandas.concat([header, tmp]))
                                        if remainder > 0 :
                                            tmp = whole_df.copy().truncate(before = divider * pivot, after = len(whole_df.index))
                                            tmp.columns = range(tmp.shape[1])
                                            df_list.append(pandas.concat([header, tmp]))
                                    else:
                                        if remainder == 0:
                                            tmp = whole_df.copy().truncate(after = pivot -1)
                                            tmp.columns = range(tmp.shape[1])
                                            df_list.append(pandas.concat([header, tmp]))
                                            tmp = whole_df.copy().truncate(before = pivot)
                                            tmp.columns = range(tmp.shape[1])
                                            df_list.append(pandas.concat([header, tmp]))
                                        else:
                                            tmp = whole_df.copy().truncate(after = pivot)
                                            tmp.columns = range(tmp.shape[1])
                                            df_list.append(pandas.concat([header, tmp]))
                                            tmp = whole_df.copy().truncate(before = pivot+1)
                                            tmp.columns = range(tmp.shape[1])
                                            df_list.append(pandas.concat([header, tmp]))


                                    first_df = df_list[0]

                                    print(f"first_df")
                                    print(first_df)

                                    table_shape = new_slide.Shapes.AddTable(NumRows = len(first_df.index), NumColumns = len(first_df.columns)
                                                                            # ,
                                                                            # Left = placeholder_info["content"][0]["Left"],
                                                                            # Top = placeholder_info["content"][0]["Top"],
                                                                            # Width = placeholder_info["content"][0]["Width"],
                                                                            # Height = placeholder_info["content"][0]["Height"]
                                                                            )
                                    table = table_shape.Table
                                    # print(f"firstRow: {first_df.iloc[0]}")
                                    for k in range(len(first_df.index)):
                                        for l in range(len(first_df.columns)):
                                            # cell = table.Cell(k+1, l+1)
                                            table.Cell(k+1, l+1).Shape.TextFrame.TextRange.Text = str(first_df.iloc[k,l])

                                    this_height = new_slide.Shapes.Item(index1+1).Height
                                    print(f"height: {this_height}")
                                    if this_height <= placeholder_info["content"][0]["Height"]:
                                        for j in range(1, divider):
                                            new_slide.Copy()
                                            presentation.Slides.Paste(presentation.Slides.Count + 1)
                                            new_slide = presentation.Slides(presentation.Slides.Count)
                                            for element in range(new_slide.Shapes.Count):
                                                t_shape = new_slide.Shapes.Item(element+1)
                                                if t_shape.HasTable:
                                                    t_shape.Delete()
                                                    this_df = df_list[j]
                                                    table_shape = new_slide.Shapes.AddTable(NumRows = len(this_df.index), NumColumns = len(this_df.columns)
                                                                                            # ,
                                                                                            # Left = placeholder_info["content"][0]["Left"],
                                                                                            # Top = placeholder_info["content"][0]["Top"],
                                                                                            # Width = placeholder_info["content"][0]["Width"],
                                                                                            # Height = placeholder_info["content"][0]["Height"]
                                                                                            )
                                                    table = table_shape.Table
                                                    for m in range(len(this_df.index)):
                                                        for n in range(len(this_df.columns)):
                                                            cell = table.Cell(m+1, n+1)
                                                            cell.Shape.TextFrame.TextRange.Text = str(this_df.iloc[m,n])

                                    else: divider = divider +1
                            else:
                                header.columns = range(header.shape[1])
                                csv_df.columns = range(csv_df.shape[1])
                                csv_df = pandas.concat([header, csv_df])
                                for k in range(len(first_df.index)):
                                    for l in range(len(first_df.columns)):
                                        # cell = table.Cell(k+1, l+1)
                                        table.Cell(k+1, l+1).Shape.TextFrame.TextRange.Text = str(csv_df.iloc[k,l])

                        elif (content["type"] == "image"):
                            new_slide.Shapes.AddPicture(FileName = placeholder_info["content"][index2]["item"],
                                                                    LinkToFile = False,
                                                                    SaveWithDocument = True,
                                                                    Left = placeholder_info["content"][index2]["Left"],
                                                                    Top = placeholder_info["content"][index2]["Top"],
                                                                    Width = placeholder_info["content"][index2]["Width"],
                                                                    Height = placeholder_info["content"][index2]["Height"])

        # presentation.SaveAS(name + '.pptx')
        # presentation.Close()

        # PPTApp.Quit()
                                    # layout = presentation.Designs.Item(1).SlideMaster.CustomLayouts.Item(slide_layout[slide_info["layout"]])
                                    # new_slide = presentation.Slides.AddSlide(presentation.Slides.Count + 1, layout)



                    

        


        presentation.SaveAS(name + '.pptx')
        presentation.Close()

        PPTApp.Quit()

    generate_slide(paper_summary)

    # """
    # Generate slide with given paper summary information.

    # :param paper_summary: summarized paper contents
    # (The following format is just an example. You can define better format.)
    # (e.g. [
    #     {
    #         'title': 'Introduction',
    #         'summary': ['This paper is ...', 'We discovered that ...', ...],
    #         'figure': ['/path/to/file', '/path/to/file', ...]
    #     },
    #     {
    #         'title': 'Overview of CFS and ULE',
    #         'subtitle': 'Linux CFS',
    #         'summary': ['Per-core scheduling: Linuxs CFS implements ...', ...],
    #         'figure': ['/path/to/file', ...]
    #     }
    # ])
    # :return: pptx slide file
    # """
    # paper_summary = json.loads(paper_summary)   #문자열 paper_summary를 json객체로 변경
    # prs = Presentation()                        #새 슬라이드 생성

    # """
    # 일단 slide_info의 가장 기본형으로 가정하고 개발
    # slide_info[0]: Title 페이지
    # slide_info[1]: Title and content
    # slide_info[2~n-1]: 부제, 내용
    # slide_info[n]: Q & A page
    # """
    # name = paper_summary['user_name']
    # for slide_info in paper_summary:                 #slide_info에는 각 slide페이지에 대한 정보가 담겨져 있음
    #     slide = prs.slides.add_slide(slide_info['layout'])
    #     """
    #     slide layout
    #     0: Title
    #     1: Title and Content
    #     2: Section Header (sometimes called Segue)
    #     3: Two Content (side by side bullet textboxes)
    #     4: Comparison (same but additional title for each side by side content box)
    #     5: Title Only
    #     6: Blank
    #     7: Content with Caption
    #     8: Picture with Caption
    #     """
    #     if slide_info['layout'] == 0:
    #         title = slide.shapes.title
    #         subtitle = slide.placehoders[1]

    #         title.text = slide_info['contents']['title']
    #         subtitle.text = slide_info['contents']['title']

    # prs.save(name+'.pptx')
