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
import os
import fitz
import cv2

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
        if request.POST.get('selected'):
            template_id = request.POST.get('selected')
            template = settings.BASE_DIR / \
                f'static/common/potx/template{template_id}.potx'
            usertemplate = False
        else:
            form = FileUploadForm(request.POST, request.FILES)
            if form.is_valid():
                file = request.FILES['file']
                fs = FileSystemStorage()
                potx_file = fs.save(file.name, file)
                template = settings.MEDIA_ROOT / potx_file
            usertemplate = True

        name, _ = os.path.splitext(summary_json_file)
        option = {
            'title': name,
            'username': 'username',
            'titlefont': 'Arial',
            'subtitlefont': 'Arial',
            'font': 'Arial',
            'spacing': 35,
            'wide': True,
            'usertemplate': usertemplate,
            'template': str(template)
        }
        with open(settings.MEDIA_ROOT / f'{name}_option.json', 'w') as option_file:
            json.dump(option, option_file)
        paper_summary = settings.MEDIA_ROOT / summary_json_file
        print(f'extract_image({name}.pdf)')
        extract_image(f'{name}.pdf')
        print(f'generate_slide({paper_summary}, {template}, {option})')
        generate_slide(settings.MEDIA_ROOT / summary_json_file, settings.BASE_DIR /
                       template, option)

        return redirect('paper2slide:adjust_options', pptx_file_name=f'{name}.pptx')
    template_list = [
        {'id': i, 'name': f'Template {i}', 'thumbnail': f'template{i}.png'} for i in range(1, 11)
    ]
    form = FileUploadForm()
    return render(request, 'paper2slide/step-2.html', {'template_list': template_list, 'form': form})

# TODO: remove


def apply_template(request):
    # Create slide here
    return redirect('paper2slide:adjust_options')


def upload_template(request, summay_json_file):
    return HttpResponse("Upload template!")


def adjust_options(request, pptx_file_name):
    name, _ = os.path.splitext(pptx_file_name)
    option_file = open(settings.MEDIA_ROOT / f'{name}_option.json', 'r')
    option_json = json.load(option_file)
    option_file.close()

    if request.method == 'POST':
        form = SlideOptionForm(request.POST)
        if form.is_valid():
            new_option = {
                'title': form.cleaned_data['title'],
                'username': form.cleaned_data['username'],
                'titlefont': form.cleaned_data['titlefont'],
                'subtitlefont': form.cleaned_data['subtitlefont'],
                'font': form.cleaned_data['font'],
                'spacing': form.cleaned_data['spacing'],
                'wide': form.cleaned_data['wide'] == 'True',
                'usertemplate': option_json['usertemplate'],
                'template': option_json['template']
            }
            with open(settings.MEDIA_ROOT / f'{name}_option.json', 'w') as file:
                json.dump(new_option, file)

            template = new_option['template']
            print(
                f'generate_slide({settings.MEDIA_ROOT / name}.json, {template}, {new_option})')
            generate_slide(settings.MEDIA_ROOT /
                           f'{name}.json', template, new_option)
            pythoncom.CoInitialize()
            powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
            powerpoint.Visible = True
            deck = powerpoint.Presentations.Open(
                settings.MEDIA_ROOT / pptx_file_name)
            deck.SaveAs(settings.MEDIA_ROOT / f'{name}_preview.pdf', 32)
            deck.Close()
            powerpoint.Quit()
            return render(request, 'paper2slide/step-3.html',
                          {'pdf_file_name': f'{name}_preview.pdf',
                           'form': form})

    form = SlideOptionForm(initial={
        'title': option_json['title'],
        'username': option_json['username'],
        'titlefont': option_json['titlefont'],
        'subtitlefont': option_json['subtitlefont'],
        'font': option_json['font'],
        'spacing': option_json['spacing'],
        'wide': option_json['wide']
    })
    pythoncom.CoInitialize()
    powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
    powerpoint.Visible = True
    deck = powerpoint.Presentations.Open(settings.MEDIA_ROOT / pptx_file_name)
    deck.SaveAs(settings.MEDIA_ROOT / f'{name}_preview.pdf', 32)
    deck.Close()
    powerpoint.Quit()
    return render(request, 'paper2slide/step-3.html', {'pdf_file_name': f'{name}_preview.pdf', 'form': form})

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
                output.append({"role": "sectionHeading",
                               "content": "REFERENCES"})
                reference_flag = 1
            elif content["role"] == "sectionHeading" or content["role"] == None:
                if content["content"].upper().strip() == "ACKNOWLEDGEMENTS":
                    break
                elif content["content"].upper().strip() == "ABSTRACT":
                    output.append({"role": "sectionHeading",
                                   "content": "ABSTRACT"})
                elif len(output) > 0 and content["role"] == None and output[-1]["role"] == None:
                    if len(content["content"]) > 10:
                        output[-1]["content"] = output[-1]["content"] + \
                            " "+content["content"]
                else:
                    if len(output) > 0 and (not content["content"][0].isdigit()) and output[-1]["role"] == None:
                        output[-1]["content"] = output[-1]["content"] + \
                            " "+content["content"]

                    else:
                        output.append(
                            {"role": content["role"], "content": content["content"]})
        elif reference_flag == 1:
            if output[-1]["content"] == "REFERENCES":
                output.append(
                    {"role": content["role"], "content": content["content"]})
            elif content["role"] == "sectionHeading":
                reference_flag == 0
            else:
                output[-1]["content"] = output[-1]["content"] + \
                    " "+content["content"]

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
                processed_data["sentences"][i]["tables"] = find_pattern_match_position(
                    processed_data["sentences"][i]["content"], processed_data["sentences"][i]["summarized"], "table")
            else:
                processed_data["sentences"][i]["tables"] = []
        else:
            processed_data["sentences"][i]["tables"] = []
    for i in range(len(processed_data["sentences"])):
        if 'summarized' in processed_data["sentences"][i]:
            if processed_data["sentences"][i]["content"] is not None:
                processed_data["sentences"][i]["figures"] = find_pattern_match_position(
                    processed_data["sentences"][i]["content"], processed_data["sentences"][i]["summarized"], 'figure')
            else:
                processed_data["sentences"][i]["figures"] = []
        else:
            processed_data["sentences"][i]["figures"] = []

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


def generate_slide(paper_summary, template, option):
    slide_layout = {"title": 1,
                    "title and Content": 2,
                    "Section Header": 3,
                    "Two Content": 4,
                    "Comparison": 5,
                    "title Only": 6,
                    "Blank": 7,
                    "Content with Caption": 8,
                    "Picture with Caption": 9}

    current_dir = os.getcwd()
    now_dir = current_dir
   # parent_dir = os.path.dirname(current_dir)
    uploads_dir = os.path.join(current_dir, "uploads")
    table_dir = os.path.dirname(uploads_dir)
    picture_dir = os.path.join(uploads_dir, "cropped")
    original_file_name = os.path.basename(paper_summary)

    def just_insert_text(presentation, title, summary_seq, option):

        layout = presentation.Designs.Item(1).SlideMaster.CustomLayouts.Item(
            slide_layout["title and Content"])
        new_slide = presentation.Slides.AddSlide(
            presentation.Slides.Count+1, layout)

        new_slide.Select()
        title = title

        # 일단 첫 페이지 만들어보기
        new_slide.Shapes.Item(1).TextFrame.TextRange.Text = title
        new_slide.Shapes.Item(
            1).TextFrame.TextRange.Font.Name = option["subtitlefont"]

        if (len(summary_seq) == 1):
            new_slide.Shapes.Item(2).TextFrame.TextRange.Text = summary_seq[0]
            new_slide.Shapes.Item(
                2).TextFrame.TextRange.ParagraphFormat.Bullet.Visible = True
            new_slide.Shapes.Item(
                2).TextFrame.TextRange.ParagraphFormat.LineRuleWithin = False
            new_slide.Shapes.Item(
                2).TextFrame.TextRange.ParagraphFormat.SpaceWithin = option["spacing"]
            new_slide.Shapes.Item(
                2).TextFrame.TextRange.Font.Name = option["font"]
            return

        body = new_slide.Shapes.Item(2).TextFrame.TextRange
        Textinsertion = 0
        for numinput, _ in enumerate(summary_seq):
            numiput = numinput+1

        for index1, _ in enumerate(summary_seq):
            this_paragraph = body.Paragraphs(index1+1)
            if (numinput == Textinsertion + 1):
                this_paragraph.Text = summary_seq[index1]
                Textinsertion = Textinsertion + 1
            else:
                this_paragraph.Text = summary_seq[index1] + "\n"
                Textinsertion = Textinsertion + 1
            this_paragraph.ParagraphFormat.Bullet.Visible = True

        new_slide.Shapes.Item(
            2).TextFrame.TextRange.ParagraphFormat.LineRuleWithin = False
        new_slide.Shapes.Item(
            2).TextFrame.TextRange.ParagraphFormat.SpaceWithin = option["spacing"]
        new_slide.Shapes.Item(2).TextFrame.TextRange.Font.Name = option["font"]

        # 만약 글자가 너무 작으면...
        n = 2
        while (new_slide.Shapes.Item(2).TextFrame.TextRange.Font.Size < 26):
            # 현재 페이지 삭제
            for i in range(n-1):
                presentation.Slides(presentation.Slides.Count).Delete()

            # 일단 입력할 문장들을 n개로 분할하기
            summary_seq_seq = []
            sub_list_size = len(summary_seq) // n
            start = 0

            for _ in range(n-1):
                sublist = summary_seq[start:start+sub_list_size]
                summary_seq_seq.append(sublist)
                start += sub_list_size

            sublist = summary_seq[start:]
            summary_seq_seq.append(sublist)

            # 입력하기
            for _, seq in enumerate(summary_seq_seq):
                layout = presentation.Designs.Item(1).SlideMaster.CustomLayouts.Item(
                    slide_layout["title and Content"])
                new_slide = presentation.Slides.AddSlide(
                    presentation.Slides.Count+1, layout)
                new_slide.Shapes.Item(1).TextFrame.TextRange.Text = title
                new_slide.Shapes.Item(
                    1).TextFrame.TextRange.Font.Name = option["subtitlefont"]
                body = new_slide.Shapes.Item(2).TextFrame.TextRange
                Textinsertion = 0

                for numinput, _ in enumerate(seq):
                    numinput = numinput+1

                for index1, _ in enumerate(seq):
                    this_paragraph = body.Paragraphs(index1+1)
                    if (numinput == Textinsertion + 1):
                        this_paragraph.Text = seq[index1]
                        Textinsertion = Textinsertion + 1
                    else:
                        this_paragraph.Text = seq[index1] + "\n"
                        Textinsertion = Textinsertion + 1
                    this_paragraph.ParagraphFormat.Bullet.Visible = True

                new_slide.Shapes.Item(
                    2).TextFrame.TextRange.ParagraphFormat.LineRuleWithin = False
                new_slide.Shapes.Item(
                    2).TextFrame.TextRange.ParagraphFormat.SpaceWithin = option["spacing"]
                new_slide.Shapes.Item(
                    2).TextFrame.TextRange.Font.Name = option["font"]

            n += 1

    def insert_text_with_picture(presentation, title, summary_seq, picture_seq, option):

        if type(picture_seq[0]) == list:
            temp = []
            for pic_set in picture_seq:
                temp.append(pic_set[0])
                temp.append(pic_set[1])
            picture_seq = temp

        index_of_picture_sentence = []
        picture_list = []
        for i, val in enumerate(picture_seq):
            if i & 1 == True:
                index_of_picture_sentence.append(val)
            else:
                picture_list.append(val)

        print(f"picture_list: {picture_list}")
        print(f"index_of_picture_sentence: {index_of_picture_sentence}")

        summary_seq_seq = []
        LastSentenceDoesntIncludePictureFlag = True
        start = 0
        pivot = 0
        for index, pivot in enumerate(index_of_picture_sentence):
            print(f"pivot: {pivot}")

            # if len(picture_seq) == 2:
            #     seq = summary_seq[0:pivot-1]
            #     seq.append(None)
            #     summary_seq_seq.append(seq)
            #     seq = []
            #     seq.append(summary_seq[pivot])
            #     seq.append(picture_list[0])
            #     summary_seq_seq.append(seq)
            #     pass

            if pivot == 0:
                seq = []
                seq.append(summary_seq[pivot])
                seq.append(picture_list[index])
                summary_seq_seq.append(seq)
                start = 1
            elif pivot == len(summary_seq):
                LastSentenceDoesntIncludePictureFlag = False
                seq = summary_seq[start:pivot]
                seq.append(None)
                summary_seq_seq.append(seq)
                seq = []
                seq.append(summary_seq[pivot])
                seq.append(picture_list[index])
                summary_seq_seq.append(seq)
                start = pivot + 1
            else:
                seq = summary_seq[start:pivot]
                seq.append(None)
                summary_seq_seq.append(seq)
                seq = []
                seq.append(summary_seq[pivot])
                seq.append(picture_list[index])
                summary_seq_seq.append(seq)
                start = pivot + 1

        if LastSentenceDoesntIncludePictureFlag:
            seq = summary_seq[pivot+1:]
            seq.append(None)
            summary_seq_seq.append(seq)

        for i, s in enumerate(summary_seq_seq):
            print(f"summary_seq_seq[{i}]: {s}")

        # 이제 만들어놓은 summary_seq_seq을 바탕으로 슬라이드 생성!
        for sentences in summary_seq_seq:
            if sentences[-1] == None:
                if len(sentences) == 1:
                    pass
                else:
                    print(
                        f"sentences[0:len(sentences)-1]: {sentences[0:len(sentences)-1]}")
                    just_insert_text(presentation, title,
                                     sentences[0:len(sentences)-1], option)
            else:
                layout = presentation.Designs.Item(1).SlideMaster.CustomLayouts.Item(
                    slide_layout["Picture with Caption"])
                new_slide = presentation.Slides.AddSlide(
                    presentation.Slides.Count+1, layout)

                new_slide.Shapes.Item(1).TextFrame.TextRange.Text = title
                new_slide.Shapes.Item(
                    1).TextFrame.TextRange.Font.Name = option["subtitlefont"]

                print(
                    "-------------------------------------------------------------------")
                print(f"picture_dir: {picture_dir}")
                print(
                    "-------------------------------------------------------------------")

                new_slide.Shapes.Item(2).Fill.UserPicture(
                    os.path.join(picture_dir, f"{sentences[1]}.png"))

                new_slide.Shapes.Item(
                    3).TextFrame.TextRange.Text = sentences[0]
                new_slide.Shapes.Item(
                    3).TextFrame.TextRange.ParagraphFormat.Bullet.Visible = False
                new_slide.Shapes.Item(
                    3).TextFrame.TextRange.ParagraphFormat.LineRuleWithin = False
                new_slide.Shapes.Item(
                    3).TextFrame.TextRange.ParagraphFormat.SpaceWithin = option["spacing"]
                new_slide.Shapes.Item(
                    3).TextFrame.TextRange.Font.Name = option["font"]

    def insert_text_with_table(presentation, title, summary_seq, table_seq, option):

        if type(table_seq[0]) == list:
            temp = []
            for table_set in table_seq:
                temp.append(table_set[0])
                temp.append(table_set[1])
            table_seq = temp

        index_of_table_sentence = []
        table_list = []
        for i, val in enumerate(table_seq):
            if i & 1 == True:
                index_of_table_sentence.append(val)
            else:
                table_list.append(val)

        print(f"table_list: {table_list}")
        print(f"index_of_table_sentence: {index_of_table_sentence}")

        summary_seq_seq = []
        LastSentenceDoesntIncludeTableFlag = True
        start = 0
        pivot = 0
        for index, pivot in enumerate(index_of_table_sentence):
            print(f"pivot: {pivot}")
            print(f'len(summary_seq): {len(summary_seq)}')

            # if len(table_seq) == 2:
            #     seq = summary_seq[0:pivot]
            #     seq.append(None)
            #     summary_seq_seq.append(seq)
            #     seq = []
            #     seq.append(summary_seq[pivot])
            #     seq.append(table_list[0])
            #     summary_seq_seq.append(seq)
            #     pass

            if pivot == 0:
                seq = []
                seq.append(summary_seq[pivot])
                seq.append(table_list[index])
                summary_seq_seq.append(seq)
                start = 1
            elif pivot == len(summary_seq):
                LastSentenceDoesntIncludeTableFlag = False
                seq = summary_seq[start:pivot]
                seq.append(None)
                summary_seq_seq.append(seq)
                seq = []
                seq.append(summary_seq[pivot])
                seq.append(table_list[index])
                summary_seq_seq.append(seq)
                start = pivot + 1
            else:
                seq = summary_seq[start:pivot]
                seq.append(None)
                summary_seq_seq.append(seq)
                seq = []
                seq.append(summary_seq[pivot])
                seq.append(table_list[index])
                summary_seq_seq.append(seq)
                start = pivot + 1

        if LastSentenceDoesntIncludeTableFlag:
            seq = summary_seq[pivot+1:]
            seq.append(None)
            summary_seq_seq.append(seq)

        for i, s in enumerate(summary_seq_seq):
            print(f"summary_seq_seq[{i}]: {s}")

        # #임시로 하나 만들어보기
        # layout = presentation.Designs.Item(1).SlideMaster.CustomLayouts.Item(slide_layout["title and Content"])
        # new_slide = presentation.Slides.AddSlide(presentation.Slides.Count+1, layout)

        # new_slide.Shapes.Item(1).TextFrame.TextRange.Text = option["title"]
        # new_slide.Shapes.Item(1).TextFrame.TextRange.Font.Name = option["titlefont"]

        # excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
        # workbook = excel.Workbooks.Open(now_dir + "\\" + table_list[0])
        # worksheet = workbook.Worksheets('Sheet')
        # base_cell = worksheet.Range("A1")
        # table_range = base_cell.CurrentRegion
        # table_range.Copy()

        # new_slide.Shapes.PasteSpecial()

        # 이제 만들어놓은 summary_seq_seq을 바탕으로 슬라이드 생성!
        for sentences in summary_seq_seq:
            if sentences[-1] == None:
                if len(sentences) == 1:
                    pass
                else:
                    print(
                        f"sentences[0:len(sentences)-1]: {sentences[0:len(sentences)-1]}")
                    just_insert_text(presentation, title,
                                     sentences[0:len(sentences)-1], option)
            else:
                layout = presentation.Designs.Item(1).SlideMaster.CustomLayouts.Item(
                    slide_layout["title and Content"])
                new_slide = presentation.Slides.AddSlide(
                    presentation.Slides.Count+1, layout)

                new_slide.Shapes.Item(1).TextFrame.TextRange.Text = title
                new_slide.Shapes.Item(
                    1).TextFrame.TextRange.Font.Name = option["subtitlefont"]

                pythoncom.CoInitialize()
                excel = win32com.client.gencache.EnsureDispatch(
                    'Excel.Application')
                workbook = excel.Workbooks.Open(
                    table_dir + "\\" + sentences[-1] + ".xlsx")
                print(
                    "--------------------------------------------------------------------------")
                print(table_dir + "\\" + sentences[-1] + ".xlsx")
                print(
                    "--------------------------------------------------------------------------")
                worksheet = workbook.ActiveSheet
                # table = worksheet.ListObjects(1)
                # table_range = table.Range()
                # table_data = table_range.Value
                base_cell = worksheet.Range("A1")
                table_range = base_cell.CurrentRegion
                table_range.Copy()
                table_range = new_slide.Shapes.PasteSpecial()
                excel.Quit()
                # pythoncom.CoUninitialize()

                # table = new_slide.Shapes.Item(2).Table
                # print("--------------------------------------------------------------------------")
                # print(type(table))
                # print("--------------------------------------------------------------------------")
                n = 0
                for i, shape in enumerate(new_slide.Shapes):
                    print(
                        f"i: {i} // shape: {shape} // HasTable: {shape.HasTable}")
                    if shape.HasTable:
                        table = shape
                        print(
                            "--------------------------------------------------------------------------")
                        print("TABLE!!!!")
                        print(
                            "--------------------------------------------------------------------------")
                        n = i+1
                        break

                # table = new_slide.Shapes.Item(n)
                # table.Table.ScaleProportionally(min(1000//table.Width, 700//table.Height))
                # table.Table.title = sentences[1]
                scale = min(1000//table_range.Width, 700//table_range.Height)
                table_range.ScaleWidth(scale, 0)
                table_range.ScaleHeight(scale, 0)
                # table_range.Item().Table.title = sentences[1]

                new_slide.Shapes.AddTextbox(1, 100, 100, 100, 100)
                new_slide.Shapes.Item(
                    3).TextFrame.TextRange.Text = sentences[0]
                new_slide.Shapes.Item(
                    3).TextFrame.TextRange.ParagraphFormat.Bullet.Visible = False
                new_slide.Shapes.Item(
                    3).TextFrame.TextRange.ParagraphFormat.LineRuleWithin = False
                new_slide.Shapes.Item(
                    3).TextFrame.TextRange.ParagraphFormat.SpaceWithin = option["spacing"]
                new_slide.Shapes.Item(
                    3).TextFrame.TextRange.Font.Name = option["font"]
                new_slide.Shapes.Item(3).TextFrame.WordWrap = False

                # new_slide.Shapes.Item(3).Top = table_shape_range.Top - new_slide.Shapes.Item(3).Height
                new_slide.Shapes.Item(3).Top = (
                    table_range.Top + table_range.Height + new_slide.Shapes.Item(3).Height)
                new_slide.Shapes.Item(3).Left = table_range.Left

                # table.Name = "a"
                new_slide.Shapes.Item(3).Name = "b"

                # table.Left = (new_slide.Master.Width / 2) - (table.Width / 2)
                # group = new_slide.Shapes.Range(["a", "b"]).Group()
                # table_shape_range.Align(1, True)    #https://learn.microsoft.com/en-us/office/vba/api/office.msoaligncmd
                # new_slide.Shapes.Range(["a","b"]).Align(1, True)

                # for row in table.Table.Rows:
                #     for cell in row.Cells:
                #         cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = win32com.client.constants.ppAlignCenter

                # group.Align(4, True)

    def insert_text_with_both(presentation, title, summary_seq, figure_seq, table_seq, option):

        if type(figure_seq[0]) == list:
            temp = []
            for pic_set in figure_seq:
                temp.append(pic_set[0])
                temp.append(pic_set[1])
            figure_seq = temp

        if type(table_seq[0]) == list:
            temp = []
            for table_set in table_seq:
                temp.append(table_set[0])
                temp.append(table_set[1])
            table_seq = temp

        index_of_contents_sentence = []
        index_of_table_sentence = []
        index_of_picture_sentence = []
        picture_list = []
        table_list = []
        contents_list = []

        for i, val in enumerate(table_seq):
            if i & 1 == True:
                index_of_table_sentence.append(val)
            else:
                table_list.append(val)

        print(f"table_list: {table_list}")
        print(f"index_of_table_sentence: {index_of_table_sentence}")

        for i, val in enumerate(figure_seq):
            if i & 1 == True:
                index_of_picture_sentence.append(val)
            else:
                picture_list.append(val)

        print(f"picture_list: {picture_list}")
        print(f"index_of_picture_sentence: {index_of_picture_sentence}")

        index_of_contents_sentence = index_of_picture_sentence + index_of_table_sentence
        index_of_contents_sentence = sorted(index_of_contents_sentence)

        for index in index_of_contents_sentence:
            try:
                i = figure_seq.index(index)
                contents_list.append(figure_seq[i-1])
            except ValueError:
                contents_list.append(table_seq[table_seq.index(index)-1])

        summary_seq_seq = []
        LastSentenceDoesntIncludeTableFlag = True
        LastSentenceDoesntIncludePictureFlag = True

        start = 0
        pivot = 0
        for index, pivot in enumerate(index_of_contents_sentence):
            print(f"pivot: {pivot}")
            print(f'len(summary_seq): {len(summary_seq)}')

            # if len(table_seq) == 2:
            #     seq = summary_seq[0:pivot]
            #     seq.append(None)
            #     summary_seq_seq.append(seq)
            #     seq = []
            #     seq.append(summary_seq[pivot])
            #     seq.append(table_list[0])
            #     summary_seq_seq.append(seq)
            #     pass
            if pivot == 0:
                seq = []
                seq.append(summary_seq[pivot])
                seq.append(contents_list[index])
                summary_seq_seq.append(seq)
                start = 1
            elif pivot == len(summary_seq):
                LastSentenceDoesntIncludeTableFlag = False
                seq = summary_seq[start:pivot]
                seq.append(None)
                summary_seq_seq.append(seq)
                seq = []
                seq.append(summary_seq[pivot])
                seq.append(contents_list[index])
                summary_seq_seq.append(seq)
                start = pivot + 1
            else:
                seq = summary_seq[start:pivot]
                seq.append(None)
                summary_seq_seq.append(seq)
                seq = []
                seq.append(summary_seq[pivot])
                seq.append(contents_list[index])
                summary_seq_seq.append(seq)
                start = pivot + 1

        if LastSentenceDoesntIncludeTableFlag and LastSentenceDoesntIncludePictureFlag:
            seq = summary_seq[pivot+1:]
            seq.append(None)
            summary_seq_seq.append(seq)

        for i, s in enumerate(summary_seq_seq):
            print(f"summary_seq_seq[{i}]: {s}")

        for sentences in summary_seq_seq:
            if len(sentences) == 1:
                pass
            elif sentences[-1] == None:
                just_insert_text(presentation, title,
                                 sentences[0: len(sentences) - 1], option)
            # elif sentences[-1].split(".")[1] == "png":
            elif sentences[-1].lower().startswith('fig'):
                tmp = []
                tmp.append(sentences[-1])
                tmp.append(0)
                insert_text_with_picture(
                    presentation, title, sentences[0:1], tmp, option)
            # elif sentences[-1].split(".")[1] == "xlsx":
            elif sentences[-1].lower().startswith('table'):
                tmp = []
                tmp.append(sentences[-1])
                tmp.append(0)
                insert_text_with_table(
                    presentation, title, sentences[0:1], tmp, option)

    ##########################################################################################################################################

    def generate_slides(paper_summary, template, option):

        with open(paper_summary, encoding='utf-8') as file:
            paper_summary = json.load(file)

        # with open(option, encoding='utf-8') as file:
        #     option = json.load(file)

        paper_summary = paper_summary["sentences"]

        print(paper_summary)

        # for subsection in paper_summary:
        #     print(subsection)
        #     print("-----------------------------------")

        for subsection in paper_summary:
            if subsection["summarized"] is not None:
                for sentence in subsection["summarized"]:
                    with open("tmp.txt", "a", encoding='utf-8') as file:
                        file.write(str(sentence))

        save_name = str(original_file_name.split(".")[0]) + ".pptx"

        PPTApp = win32com.client.gencache.EnsureDispatch(
            "PowerPoint.Application")
        presentation = PPTApp.Presentations.Add()

        # if(template != "basic"):
        #     presentation.ApplyTemplate(now_dir+ "\\" +template)

        # 표지 만들기
        layout = presentation.Designs.Item(
            1).SlideMaster.CustomLayouts.Item(slide_layout["title"])
        new_slide = presentation.Slides.AddSlide(
            presentation.Slides.Count+1, layout)

        new_slide.Shapes.Item(1).TextFrame.TextRange.Text = option["title"]
        new_slide.Shapes.Item(
            1).TextFrame.TextRange.Font.Name = option["titlefont"]
        new_slide.Shapes.Item(2).TextFrame.TextRange.Text = option["username"]
        new_slide.Shapes.Item(
            2).TextFrame.TextRange.Font.Name = option["titlefont"]

        # #목차 만들기

        subtitles = []

        for sub in paper_summary:
            subtitles.append(sub["title"])

        just_insert_text(presentation, "Contents", subtitles, option)

        # layout = presentation.Designs.Item(1).SlideMaster.CustomLayouts.Item(slide_layout["title and Content"])
        # new_slide = presentation.Slides.AddSlide(presentation.Slides.Count+1, layout)

        # new_slide.Shapes.Item(1).TextFrame.TextRange.Text = "Contents"
        # new_slide.Shapes.Item(1).TextFrame.TextRange.Font.Name = option["titlefont"]
        # body = new_slide.Shapes.Item(2).TextFrame.TextRange
        # Textinsertion = 0
        # for numheader, _  in enumerate(paper_summary):
        #     numheader = numheader+1

        # for index1, _ in enumerate(paper_summary):
        #     this_paragraph = body.Paragraphs(index1+1)
        #     if (numheader == Textinsertion +1):
        #         this_paragraph.Text = paper_summary[index1]["title"]
        #         this_paragraph.Font.Name = option["font"]
        #         Textinsertion = Textinsertion +1
        #     else:
        #         this_paragraph.Text = paper_summary[index1]["title"] + "\n"
        #         this_paragraph.Font.Name = option["font"]
        #         Textinsertion = Textinsertion +1
        #     this_paragraph.ParagraphFormat.Bullet.Visible = True

    #######################################################################################################################################
        # 본격적인 내용
        for i, summary in enumerate(paper_summary):

            if summary["summarized"] is not None:
                # 서브섹션 타이틀 페이지 만들기
                layout = layout = presentation.Designs.Item(
                    1).SlideMaster.CustomLayouts.Item(slide_layout["Section Header"])
                new_slide = presentation.Slides.AddSlide(
                    presentation.Slides.Count+1, layout)
                new_slide.Select()

                new_slide.Shapes.Item(
                    1).TextFrame.TextRange.Text = summary["title"]
                new_slide.Shapes.Item(
                    1).TextFrame.TextRange.Font.Name = option["subtitlefont"]
                new_slide.Shapes.Item(
                    2).TextFrame.TextRange.Text = option["title"]
                new_slide.Shapes.Item(
                    2).TextFrame.TextRange.Font.Name = option["titlefont"]

                # figure와 table 개수 세기
                # n_figure = 0
                # n_table = 0
                # for n_figure, _ in enumerate(summary["figures"]):
                #     pass
                # for n_table, _ in enumerate(summary["tables"]):
                #     pass

                # n_figure = True
                # n_table = True
                # if summary["figures"] is None:
                #     n_figure = False
                # if summary["tables"] is None:
                #     n_table = False
                len_figure = len(summary["figures"])
                len_table = len(summary["tables"])

                # figure와 table이 없는 경우(text만 있는 경우) slide 만들기
                if (len_figure == 0) and (len_table == 0):
                    just_insert_text(
                        presentation, summary["title"], summary["summarized"], option)
                elif (len_figure != 0) and (len_table == 0):
                    insert_text_with_picture(
                        presentation, summary["title"], summary["summarized"], summary["figures"], option)
                elif (len_figure == 0) and (len_table != 0):
                    insert_text_with_table(
                        presentation, summary["title"], summary["summarized"], summary["tables"], option)
                else:
                    insert_text_with_both(
                        presentation, summary["title"], summary["summarized"], summary["figures"], summary["tables"], option)
            else:
                pass

    #######################################################################################################################################
        # QnA 페이지 만들기
        layout = presentation.Designs.Item(
            1).SlideMaster.CustomLayouts.Item(slide_layout["Blank"])
        new_slide = presentation.Slides.AddSlide(
            presentation.Slides.Count+1, layout)
        new_slide.Select()

        text_box = new_slide.Shapes.AddTextbox(1, 100, 100, 100, 100)
        text_frame = text_box.TextFrame
        text_frame.TextRange.Text = "QnA"
        # https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppautosize
        text_frame.AutoSize = win32com.client.constants.ppAutoSizeShapeToFitText
        text_frame.WordWrap = False
        text_frame.TextRange.Font.Size = 50
        text_frame.TextRange.Font.Name = option["titlefont"]

        text_frame.TextRange.ParagraphFormat.Alignment = win32com.client.constants.ppAlignCenter

        s = new_slide.Shapes.Range()
        # https://learn.microsoft.com/en-us/office/vba/api/office.msoaligncmd
        s.Align(1, True)
        s.Align(4, True)

        # option["usertemplate"] = False

        if (template != "basic") and option["usertemplate"] == False:
            presentation.ApplyTemplate(os.path.join(
                current_dir, "static", "common", "potx", template))

        if (template != "basic") and option["usertemplate"] == True:
            presentation.ApplyTemplate(os.path.join(uploads_dir, template))

        # if (option["wide"] == True):
        #     slide_master = presentation.SlideMaster
        #     layout = None
        #     for custom_layout in slide_master.CustomLayouts:
        #         if custom_layout.Width == 1280 and custom_layout.Height == 720:
        #             layout = custom_layout
        #             break
        #     if layout is not None:
        #         for slide in presentation.Slides:
        #             slide.Layout = layout

        if (option["wide"] == True):
            print(f"option[wide]: True")
            presentation.PageSetup.SlideWidth = 33.867 * 35
            presentation.PageSetup.SlideHeight = 19.05 * 36
        else:
            print(f"option[wide]: False")
            presentation.PageSetup.SlideWidth = 25.4 * 36
            presentation.PageSetup.SlideHeight = 19.05 * 36

        presentation.SaveAs(os.path.join(uploads_dir, save_name))
        presentation.Close()
        PPTApp.Quit()

    pythoncom.CoInitialize()
    generate_slides(paper_summary, template, option)
    pythoncom.CoUninitialize()


def extract_image(paper):

    ######################################################
    # 1. 각 페이지를 이미지로 변환하기
    ######################################################

    current_dir = os.getcwd()
    uploads_dir = os.path.join(current_dir, "uploads")

    name = paper.split(".")[0]

    doc = fitz.open(uploads_dir + "\\" + paper)

    wholepage_dir = os.path.join(uploads_dir, "whole")

    if not os.path.exists(wholepage_dir):
        os.makedirs(wholepage_dir)

    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72), dpi=None,
                              colorspace=fitz.csRGB, clip=True, alpha=True, annots=True)
        pix.save(wholepage_dir + f"\\samplepdfimage-%i.png" %
                 page.number)  # save file

    ######################################################
    # 2. 각 페이지에서 이미지 crop 하기
    ######################################################

    input = name

    source_folder_path = wholepage_dir
    file_list = os.listdir(source_folder_path)

    cropped_dir = os.path.join(uploads_dir, "cropped")

    if not os.path.exists(cropped_dir):
        os.makedirs(cropped_dir)

    n = 0

    for i, file in enumerate(file_list):

        with open(source_folder_path + "\\" + file, 'rb') as image_file:
            image_data = image_file.read()

        image = cv2.imread(wholepage_dir + "\\" + file)

        if (i == 0):
            height, width = image.shape[:2]
            remove_height = int(height/4)
            image = image[remove_height:height - int(height/8), :]

        gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

        _, binary_image = cv2.threshold(
            gray_image, 100, 255, cv2.THRESH_BINARY)

        # 외곽선 검출
        contours, _ = cv2.findContours(
            binary_image, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        # 흰색 사각형 영역 검출 및 자르기
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            if w > width//6 and h > 10:  # 흰색 사각형으로 인정할 최소 너비와 높이 설정
                cropped_image = image[y-5:y+h+5, x-5:x+w+5]
                cv2.imwrite(cropped_dir + "\\" + "figure_" +
                            str(n+1) + ".png", cropped_image)
                n += 1
