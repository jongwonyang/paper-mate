from django.http import HttpResponse
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
    return HttpResponse("P2S index.")

# TODO: Heejae


def pdf_to_text(pdf_file):
    """
    Divide given paper (pdf file) into sections,
    (e.g. Intorduction, Background, Evaluation ...)
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
    pass

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
                    for index2, content in enumerate(placeholder_info["content"]):
                        if (content["type"] == "text"):
                            this_paragraph = body.Paragraphs(index2+1)
                            this_paragraph.Text = content["item"] + "\n"
                            #this_paragraph.IndentLevel = content["level"] +1
                            this_paragraph.ParagraphFormat.Bullet.Visible = True

                        elif (content["type"] == "table"):
                            table_csv_file = placeholder_info["content"][0]["item"]
                            csv_df = pandas.read_csv(table_csv_file)

                            numRow, numCol = csv_processor.RowCol(table_csv_file)
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
                                    
                                    if divider > 2:
                                        for i in range(divider):
                                                if(i * pivot < (i+1)* pivot -1):
                                                    df_list.append(whole_df.copy().truncate(before = i * pivot, after = (i+1)* pivot-1))
                                        if remainder > 0 :
                                            df_list.append(whole_df.copy().truncate(before = divider * pivot, after = len(whole_df.index)))
                                    else:
                                        if remainder == 0:
                                            df_list.append(whole_df.copy().truncate(after = pivot -1))
                                            df_list.append(whole_df.copy().truncate(before = pivot))
                                        else:
                                            df_list.append(whole_df.copy().truncate(after = pivot))
                                            df_list.append(whole_df.copy().truncate(before = pivot+1))


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
                                        for j in range(1, divider+1):
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
                
                elif (placeholder_type == "text placeholder"):
                    body = new_slide.Shapes.Item(index1+1).TextFrame.TextRange
                    for index2, content in enumerate(placeholder_info["content"]):
                        this_paragraph = body.Paragraphs(index2+1)
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

                    numRow, numCol = csv_processor.RowCol(table_csv_file)
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

                            if divider > 2:
                                for i in range(divider):
                                    if(i * pivot < (i+1)* pivot -1):    
                                        df_list.append(whole_df.copy().truncate(before = i * pivot, after = (i+1)* pivot-1))
                                if remainder > 0 :
                                    df_list.append(whole_df.copy().truncate(before = divider * pivot, after = len(whole_df.index)))
                            else:
                                if remainder == 0:
                                    df_list.append(whole_df.copy().truncate(after = pivot -1))
                                    df_list.append(whole_df.copy().truncate(before = pivot))
                                else:
                                    df_list.append(whole_df.copy().truncate(after = pivot))
                                    df_list.append(whole_df.copy().truncate(before = pivot+1))

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
                                for j in range(1, divider+1):
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
