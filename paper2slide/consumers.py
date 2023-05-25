import json, time

from paper_mate import settings
from paper2slide.views import pdf_to_text
from channels.generic.websocket import WebsocketConsumer


class PDFConsumer(WebsocketConsumer):
    def connect(self):
        self.accept()

    def disconnect(self, code):
        pass

    def receive(self, text_data):
        text_data_json = json.loads(text_data)
        filename = text_data_json['filename']
        print(f"start processing: {filename}")
        result = pdf_to_text(settings.MEDIA_ROOT / filename)
        print(result)
        self.send(text_data="Done!")