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
        print("start converting...")
        start = time.time()
        result = pdf_to_text(settings.MEDIA_ROOT / filename)
        end = time.time()
        print(result)
        print(f"time elapsed: {end - start} sec")
        name, _ = os.path.splitext(filename)
        with open(settings.MEDIA_ROOT / f'{name}.json', 'w') as json_file:
            json.dump(result, json_file)
        self.send(text_data=f'{name}.json')