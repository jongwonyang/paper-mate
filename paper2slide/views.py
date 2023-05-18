from django.http import HttpResponse
from pdfExtractor import extract_data
from summarizer import summarize_text
# TODO: Jongwon
def index(request):
    return HttpResponse("P2S index.")

# TODO: Heejae
def pdf_to_text(pdf_file):
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
    paragraphs = extract_data(pdf_file)

    output = []

    for content in paragraphs:
        if content["role"] == "sectionHeading" or content["role"] == None:
            if content["content"].lower().strip() == "acknowledgements":
                break;
            elif content["content"].upper().strip() == "ABSTRACT":
                output.append({"role":"sectionHeading",
                        "content": "ABSTRACT"})
            elif len(output)>0 and content["role"] == None and output[-1]["role"]==None:
                if len(content["content"]) > 20:
                    output[-1]["content"] = output[-1]["content"]+" "+content["content"]
                else:
                    output.append({"role":content["role"],
                            "content":content["content"]})

    output = summarize_text(output)


    return output

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