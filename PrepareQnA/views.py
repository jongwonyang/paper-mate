from django.shortcuts import render
from django.http import HttpResponse
# Create your views here.
import openai

OPENAI_API_KEY = "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"

def index(request):
    return HttpResponse("Prepare to QnA session.")


def makeQuestion(request):

    with open(request, 'r') as file:
        request = file.read()

    openai.api_key = OPENAI_API_KEY
    
    model = 'gpt-3.5-turbo'
    query = request + ' Base on last summary, make some questions to prepare QnA session with numbering. '
    messages = [
        {'role': "system", "content": "You are a helpful assistant."},
        {'role': "user", "content": query}
    ]

    response = openai.ChatCompletion.create(
        model = model,
        messages = messages
    )

    answer = response['choices'][0]['message']['content']
    
    print(answer)

