import json, time, os

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
        name, _ = os.path.splitext(filename)
        print("start converting...")
        start = time.time()
        result, _ = pdf_to_text(settings.MEDIA_ROOT / filename, settings.MEDIA_ROOT / f'{name}.json')
        end = time.time()
        print(f"time elapsed: {end - start} sec")
        self.send(text_data=f'{name}.json')