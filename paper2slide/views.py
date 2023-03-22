from django.http import HttpResponse

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
    pass