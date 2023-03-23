from django.http import HttpResponse
import json
from pptx import Presentation

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
    paper_summary = json.loads(paper_summary)   #문자열 paper_summary를 json객체로 변경
    prs = Presentation()                        #새 슬라이드 생성

    """
    일단 slide_info의 가장 기본형으로 가정하고 개발
    slide_info[0]: Title 페이지
    slide_info[1]: Title and content
    slide_info[2~n-1]: 부제, 내용
    slide_info[n]: Q & A page
    """

    for slide_info in paper_summary:                 #slide_info에는 각 slide페이지에 대한 정보가 담겨져 있음
        slide = prs.slides.add_slide(slide_info['layout'])
        """
        slide layout
        0: Title
        1: Title and Content
        2: Section Header (sometimes called Segue)
        3: Two Content (side by side bullet textboxes)
        4: Comparison (same but additional title for each side by side content box)
        5: Title Only
        6: Blank
        7: Content with Caption
        8: Picture with Caption
        """
        if slide_info['layout'] == 0:
            pass