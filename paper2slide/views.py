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
import os
import fitz
import cv2

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


def generate_slide(paper_summary, template, option):
    slide_layout = {"Title": 1,
                "Title and Content": 2,
                "Section Header": 3,
                "Two Content": 4,
                "Comparison": 5,
                "Title Only": 6,
                "Blank": 7,
                "Content with Caption": 8,
                "Picture with Caption": 9}

    now_dir = os.path.dirname(os.path.abspath(__file__))

    def just_insert_text(presentation, title, summary_seq, option):

        layout = presentation.Designs.Item(1).SlideMaster.CustomLayouts.Item(slide_layout["Title and Content"])
        new_slide = presentation.Slides.AddSlide(presentation.Slides.Count+1, layout)
        
        new_slide.Select()
        title = title

        #일단 첫 페이지 만들어보기
        new_slide.Shapes.Item(1).TextFrame.TextRange.Text = title
        new_slide.Shapes.Item(1).TextFrame.TextRange.Font.Name = option["subtitlefont"]

        if(len(summary_seq) == 1):
            new_slide.Shapes.Item(2).TextFrame.TextRange.Text = summary_seq[0]
            new_slide.Shapes.Item(2).TextFrame.TextRange.ParagraphFormat.Bullet.Visible = True
            new_slide.Shapes.Item(2).TextFrame.TextRange.ParagraphFormat.LineRuleWithin = False
            new_slide.Shapes.Item(2).TextFrame.TextRange.ParagraphFormat.SpaceWithin = option["spacing"]
            new_slide.Shapes.Item(2).TextFrame.TextRange.Font.Name = option["font"]
            return


        body = new_slide.Shapes.Item(2).TextFrame.TextRange
        Textinsertion = 0
        for numinput, _  in enumerate(summary_seq):
            numiput = numinput+1

        for index1, _ in enumerate(summary_seq):
            this_paragraph = body.Paragraphs(index1+1)
            if (numinput == Textinsertion +1):
                this_paragraph.Text = summary_seq[index1]
                Textinsertion = Textinsertion +1
            else:
                this_paragraph.Text = summary_seq[index1] + "\n"
                Textinsertion = Textinsertion +1
            this_paragraph.ParagraphFormat.Bullet.Visible = True

        new_slide.Shapes.Item(2).TextFrame.TextRange.ParagraphFormat.LineRuleWithin = False
        new_slide.Shapes.Item(2).TextFrame.TextRange.ParagraphFormat.SpaceWithin = option["spacing"]
        new_slide.Shapes.Item(2).TextFrame.TextRange.Font.Name = option["font"]

        #만약 글자가 너무 작으면...
        n = 2
        while(new_slide.Shapes.Item(2).TextFrame.TextRange.Font.Size < 26):
            #현재 페이지 삭제
            for i in range(n-1):
                presentation.Slides(presentation.Slides.Count).Delete()

            #일단 입력할 문장들을 n개로 분할하기
            summary_seq_seq = []
            sub_list_size = len(summary_seq) // n 
            start = 0
            
            for _ in range(n-1):
                sublist = summary_seq[start:start+sub_list_size]
                summary_seq_seq.append(sublist)
                start += sub_list_size
            
            sublist = summary_seq[start:]
            summary_seq_seq.append(sublist)

            #입력하기
            for _, seq in enumerate(summary_seq_seq):
                layout = presentation.Designs.Item(1).SlideMaster.CustomLayouts.Item(slide_layout["Title and Content"])
                new_slide = presentation.Slides.AddSlide(presentation.Slides.Count+1, layout)
                new_slide.Shapes.Item(1).TextFrame.TextRange.Text = title
                new_slide.Shapes.Item(1).TextFrame.TextRange.Font.Name = option["subtitlefont"]
                body = new_slide.Shapes.Item(2).TextFrame.TextRange
                Textinsertion = 0
                
                for numinput, _  in enumerate(seq):
                    numinput = numinput+1

                for index1, _ in enumerate(seq):
                    this_paragraph = body.Paragraphs(index1+1)
                    if (numinput == Textinsertion +1):
                        this_paragraph.Text = seq[index1]
                        Textinsertion = Textinsertion +1
                    else:
                        this_paragraph.Text = seq[index1] + "\n"
                        Textinsertion = Textinsertion +1
                    this_paragraph.ParagraphFormat.Bullet.Visible = True

                new_slide.Shapes.Item(2).TextFrame.TextRange.ParagraphFormat.LineRuleWithin = False
                new_slide.Shapes.Item(2).TextFrame.TextRange.ParagraphFormat.SpaceWithin = option["spacing"]
                new_slide.Shapes.Item(2).TextFrame.TextRange.Font.Name = option["font"]

            n += 1


    def insert_text_with_picture(presentation, title, summary_seq, picture_seq, option):
        
        index_of_picture_sentence = []
        picture_list =[]
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
                start = pivot +1
            else:
                seq = summary_seq[start:pivot]
                seq.append(None)
                summary_seq_seq.append(seq)
                seq = []
                seq.append(summary_seq[pivot])
                seq.append(picture_list[index])
                summary_seq_seq.append(seq)
                start = pivot +1
                
        if LastSentenceDoesntIncludePictureFlag:
            seq = summary_seq[pivot+1:]
            seq.append(None)
            summary_seq_seq.append(seq)

        for i, s in enumerate(summary_seq_seq):
            print(f"summary_seq_seq[{i}]: {s}")


        #이제 만들어놓은 summary_seq_seq을 바탕으로 슬라이드 생성!
        for sentences in summary_seq_seq:
            if sentences[-1] == None:
                if len(sentences) == 1:
                    pass
                else:
                    print(f"sentences[0:len(sentences)-1]: {sentences[0:len(sentences)-1]}")
                    just_insert_text(presentation, title, sentences[0:len(sentences)-1], option)
            else:
                layout = presentation.Designs.Item(1).SlideMaster.CustomLayouts.Item(slide_layout["Picture with Caption"])
                new_slide = presentation.Slides.AddSlide(presentation.Slides.Count+1, layout)

                new_slide.Shapes.Item(1).TextFrame.TextRange.Text = title
                new_slide.Shapes.Item(1).TextFrame.TextRange.Font.Name = option["subtitlefont"]
                
                new_slide.Shapes.Item(2).Fill.UserPicture(now_dir + "\\" + sentences[1])
                
                new_slide.Shapes.Item(3).TextFrame.TextRange.Text = sentences[0]
                new_slide.Shapes.Item(3).TextFrame.TextRange.ParagraphFormat.Bullet.Visible = False
                new_slide.Shapes.Item(3).TextFrame.TextRange.ParagraphFormat.LineRuleWithin = False
                new_slide.Shapes.Item(3).TextFrame.TextRange.ParagraphFormat.SpaceWithin = option["spacing"]
                new_slide.Shapes.Item(3).TextFrame.TextRange.Font.Name = option["font"]



    def insert_text_with_table(presentation, title, summary_seq, table_seq, option):
        index_of_table_sentence = []
        table_list =[]
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
                start = pivot +1
            else:
                seq = summary_seq[start:pivot]
                seq.append(None)
                summary_seq_seq.append(seq)
                seq = []
                seq.append(summary_seq[pivot])
                seq.append(table_list[index])
                summary_seq_seq.append(seq)
                start = pivot +1
                
        if LastSentenceDoesntIncludeTableFlag:
            seq = summary_seq[pivot+1:]
            seq.append(None)
            summary_seq_seq.append(seq)

        for i, s in enumerate(summary_seq_seq):
            print(f"summary_seq_seq[{i}]: {s}")

        # #임시로 하나 만들어보기
        # layout = presentation.Designs.Item(1).SlideMaster.CustomLayouts.Item(slide_layout["Title and Content"])
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

        #이제 만들어놓은 summary_seq_seq을 바탕으로 슬라이드 생성!
        for sentences in summary_seq_seq:
            if sentences[-1] == None:
                if len(sentences) == 1:
                    pass
                else:
                    print(f"sentences[0:len(sentences)-1]: {sentences[0:len(sentences)-1]}")
                    just_insert_text(presentation, title, sentences[0:len(sentences)-1], option)
            else:
                layout = presentation.Designs.Item(1).SlideMaster.CustomLayouts.Item(slide_layout["Title and Content"])
                new_slide = presentation.Slides.AddSlide(presentation.Slides.Count+1, layout)

                new_slide.Shapes.Item(1).TextFrame.TextRange.Text = title
                new_slide.Shapes.Item(1).TextFrame.TextRange.Font.Name = option["subtitlefont"]
                
                excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
                workbook = excel.Workbooks.Open(now_dir + "\\" + sentences[-1])
                worksheet = workbook.Worksheets('Sheet')
                base_cell = worksheet.Range("A1")
                table_range = base_cell.CurrentRegion
                table_range.Copy()


                table_shape_range = new_slide.Shapes.PasteSpecial()
                excel.Quit()

                for shape in new_slide.Shapes:
                    if shape.HasTable:
                        table = shape
                        print("TABLE!!!!")
                        break


                table.Table.ScaleProportionally(min(1000//table.Width, 700//table.Height))
                table.Table.Title = sentences[1]
                
            
                new_slide.Shapes.AddTextbox(1, 100, 100, 100, 100)
                new_slide.Shapes.Item(3).TextFrame.TextRange.Text = sentences[0]
                new_slide.Shapes.Item(3).TextFrame.TextRange.ParagraphFormat.Bullet.Visible = False
                new_slide.Shapes.Item(3).TextFrame.TextRange.ParagraphFormat.LineRuleWithin = False
                new_slide.Shapes.Item(3).TextFrame.TextRange.ParagraphFormat.SpaceWithin = option["spacing"]
                new_slide.Shapes.Item(3).TextFrame.TextRange.Font.Name = option["font"]
                new_slide.Shapes.Item(3).TextFrame.WordWrap = False

                #new_slide.Shapes.Item(3).Top = table_shape_range.Top - new_slide.Shapes.Item(3).Height
                new_slide.Shapes.Item(3).Top = (table.Top + table.Height+ new_slide.Shapes.Item(3).Height)
                new_slide.Shapes.Item(3).Left = table.Left

                table.Name = "a"
                new_slide.Shapes.Item(3).Name = "b"

                #table.Left = (new_slide.Master.Width / 2) - (table.Width / 2)
                # group = new_slide.Shapes.Range(["a", "b"]).Group()
                # table_shape_range.Align(1, True)    #https://learn.microsoft.com/en-us/office/vba/api/office.msoaligncmd
                #new_slide.Shapes.Range(["a","b"]).Align(1, True)

                for row in table.Table.Rows:
                    for cell in row.Cells:
                        cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = win32com.client.constants.ppAlignCenter
                        
                # group.Align(4, True)

    
    def insert_text_with_both(presentation, title, summary_seq, figure_seq, table_seq, option):
        index_of_contents_sentence = []
        index_of_table_sentence = []
        index_of_picture_sentence = []
        picture_list =[]
        table_list =[]
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
                start = pivot +1
            else:
                seq = summary_seq[start:pivot]
                seq.append(None)
                summary_seq_seq.append(seq)
                seq = []
                seq.append(summary_seq[pivot])
                seq.append(contents_list[index])
                summary_seq_seq.append(seq)
                start = pivot +1
                
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
                just_insert_text(presentation, title, sentences[0: len(sentences) -1], option)
            elif sentences[-1].split(".")[1] == "png":
                tmp = []
                tmp.append(sentences[-1])
                tmp.append(0)
                insert_text_with_picture(presentation, title, sentences[0:1], tmp, option)
            elif sentences[-1].split(".")[1] == "xlsx":
                tmp = []
                tmp.append(sentences[-1])
                tmp.append(0)
                insert_text_with_table(presentation, title, sentences[0:1], tmp, option)
        

    ##########################################################################################################################################
    def generate_slides(paper_summary, template, option):
        
        
        with open(paper_summary, encoding='utf-8') as file:
            paper_summary = json.load(file)

        # with open(option, encoding='utf-8') as file:
        #     option = json.load(file)

        for i in paper_summary:
            for sentence in i["summarized"]:
                with open("tmp.txt", "a") as file:
                    file.write(sentence)

        save_name = now_dir + "\\" +str(option["title"]) + ".pptx"

        PPTApp = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
        presentation = PPTApp.Presentations.Add()

        # if(template != "basic"):
        #     presentation.ApplyTemplate(now_dir+ "\\" +template)

        #표지 만들기
        layout = presentation.Designs.Item(1).SlideMaster.CustomLayouts.Item(slide_layout["Title"])
        new_slide = presentation.Slides.AddSlide(presentation.Slides.Count+1, layout)

        new_slide.Shapes.Item(1).TextFrame.TextRange.Text = option["title"]
        new_slide.Shapes.Item(1).TextFrame.TextRange.Font.Name = option["titlefont"]
        new_slide.Shapes.Item(2).TextFrame.TextRange.Text = option["username"]
        new_slide.Shapes.Item(2).TextFrame.TextRange.Font.Name = option["titlefont"]


        #목차 만들기
        layout = presentation.Designs.Item(1).SlideMaster.CustomLayouts.Item(slide_layout["Title and Content"])
        new_slide = presentation.Slides.AddSlide(presentation.Slides.Count+1, layout)

        new_slide.Shapes.Item(1).TextFrame.TextRange.Text = "Contents"
        new_slide.Shapes.Item(1).TextFrame.TextRange.Font.Name = option["titlefont"]
        body = new_slide.Shapes.Item(2).TextFrame.TextRange
        Textinsertion = 0
        for numheader, _  in enumerate(paper_summary):
            numheader = numheader+1

        for index1, _ in enumerate(paper_summary):
            this_paragraph = body.Paragraphs(index1+1)
            if (numheader == Textinsertion +1):
                this_paragraph.Text = paper_summary[index1]["Title"]
                this_paragraph.Font.Name = option["font"]
                Textinsertion = Textinsertion +1
            else:
                this_paragraph.Text = paper_summary[index1]["Title"] + "\n"
                this_paragraph.Font.Name = option["font"]
                Textinsertion = Textinsertion +1
            this_paragraph.ParagraphFormat.Bullet.Visible = True

    #######################################################################################################################################
        #본격적인 내용
        for i, summary in enumerate(paper_summary):
            
            #서브섹션 타이틀 페이지 만들기
            layout = layout = presentation.Designs.Item(1).SlideMaster.CustomLayouts.Item(slide_layout["Section Header"])
            new_slide = presentation.Slides.AddSlide(presentation.Slides.Count+1, layout)
            new_slide.Select()

            new_slide.Shapes.Item(1).TextFrame.TextRange.Text = summary["Title"]
            new_slide.Shapes.Item(1).TextFrame.TextRange.Font.Name = option["subtitlefont"]
            new_slide.Shapes.Item(2).TextFrame.TextRange.Text = option["title"]
            new_slide.Shapes.Item(2).TextFrame.TextRange.Font.Name = option["titlefont"]

            #figure와 table 개수 세기
            n_figure = 0
            n_table = 0
            for n_figure, _ in enumerate(summary["figures"]):
                pass
            for n_table, _ in enumerate(summary["tables"]):
                pass
            
            #figure와 table이 없는 경우(text만 있는 경우) slide 만들기
            if (n_figure == 0) and (n_table == 0):
                just_insert_text(presentation, summary["Title"], summary["summarized"], option)
            elif (n_figure != 0) and (n_table == 0):
                insert_text_with_picture(presentation, summary["Title"], summary["summarized"], summary["figures"],option)
            elif (n_figure == 0) and (n_table != 0):
                insert_text_with_table(presentation, summary["Title"], summary["summarized"], summary["tables"],option)
            else:
                insert_text_with_both(presentation, summary["Title"], summary["summarized"], summary["figures"], summary["tables"],option)

    #######################################################################################################################################
        #QnA 페이지 만들기
        layout = presentation.Designs.Item(1).SlideMaster.CustomLayouts.Item(slide_layout["Blank"])
        new_slide = presentation.Slides.AddSlide(presentation.Slides.Count+1, layout)
        new_slide.Select()

        text_box = new_slide.Shapes.AddTextbox(1, 100, 100, 100, 100)
        text_frame = text_box.TextFrame
        text_frame.TextRange.Text = "QnA"
        text_frame.AutoSize = win32com.client.constants.ppAutoSizeShapeToFitText  #https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppautosize
        text_frame.WordWrap = False
        text_frame.TextRange.Font.Size = 50
        text_frame.TextRange.Font.Name = option["titlefont"]

        text_frame.TextRange.ParagraphFormat.Alignment = win32com.client.constants.ppAlignCenter

        s = new_slide.Shapes.Range()
        s.Align(1, True)    #https://learn.microsoft.com/en-us/office/vba/api/office.msoaligncmd
        s.Align(4, True)

        if(option["wide"] == True):
            slide_master = presentation.SlideMaster
            layout = None
            for custom_layout in slide_master.CustomLayouts:
                if custom_layout.Width == 1280 and custom_layout.Height == 720:
                    layout = custom_layout
                    break
            if layout is not None:
                for slide in presentation.Slides:
                    slide.Layout = layout

        if(template != "basic"):
            presentation.ApplyTemplate(now_dir+ "\\" +template)

        
        presentation.SaveAs(save_name)
        presentation.Close()
        PPTApp.Quit()

    generate_slides(paper_summary, template , option)

def extract_image(paper):
    now_dir = os.getcwd()
    input = paper
    name = input.split(".")[0]

    doc = fitz.open(now_dir + "\\" + input)

    if not os.path.exists(now_dir + "\\" + name):
        os.makedirs(now_dir + "\\" + name + "_whole")

    if not os.path.exists(now_dir + "\\" + name):
        os.makedirs(now_dir + "\\" + name + "_whole2")

    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72), dpi=None,
                            colorspace=fitz.csRGB, clip=True, alpha=True, annots=True)
        pix.save(f"{name}_whole\\samplepdfimage-%i.png" % page.number)  # save file

    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72), dpi=None,
                            colorspace=fitz.csRGB, clip=False, alpha=True, annots=False)
        pix.save(f"{name}_whole2\\samplepdfimage-%i.png" % page.number)

    input = name

    source_folder_path = os.getcwd() + "\\" + input + "_whole"
    source_folder_path2 = os.getcwd() + "\\" + input + "_whole2"

    file_list = os.listdir(source_folder_path)
    file_list2 = os.listdir(source_folder_path2)

    if not os.path.exists(os.getcwd() + "\\" + input + "_cropped"):
        os.makedirs(os.getcwd() + "\\" + input + "_cropped")

    n = 0

    for i, file in enumerate(file_list):

        with open(source_folder_path+ "\\" +file, 'rb') as image_file:
            image_data = image_file.read()

        image = cv2.imread(os.getcwd()+ "\\" + input + "_whole" + "\\" + file)
        image2 = cv2.imread(os.getcwd()+ "\\" + input + "_whole2" + "\\" + file)
        
        if(i == 0):
            height, width = image.shape[:2]
            remove_height = int(height/4)
            image = image[remove_height:height - int(height/8), :]
            image2 = image2[remove_height:height - int(height/8), :]
            cv2.imwrite(os.getcwd() + "\\" + input + "_cropped" + "\\" + "fig" + str(n+1) + ".png", image)
            cv2.imshow("Cropped Image", image)
            cv2.waitKey(0)
            n += 1
            cv2.destroyAllWindows()
            

        gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

        _, binary_image = cv2.threshold(gray_image, 175, 255, cv2.THRESH_BINARY)

        # 외곽선 검출
        contours, _ = cv2.findContours(binary_image, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        # 흰색 사각형 영역 검출 및 자르기
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            if w > 10 and h > 10:  # 흰색 사각형으로 인정할 최소 너비와 높이 설정
                cropped_image = image2[y-5:y+h+5, x-5:x+w+5]
                cv2.imwrite(os.getcwd() + "\\" + input + "_cropped" + "\\" + "fig" + str(n+1) + ".png", cropped_image)
                cv2.imshow("Cropped Image", cropped_image)
                cv2.waitKey(0)
                n += 1
                cv2.destroyAllWindows()

