import json

from excelparser import ParseExcel


class ExcelToJson:

    def __init__(self, file_path, index=0, name=None, int_fields=[], float_fields=[], date_fields=[]):
        self.file_path = file_path
        self.index = index
        self.name = name
        self.int_fields = int_fields
        self.float_fields = float_fields
        self.date_fields = date_fields
        self.headers = []
        self.data = []
        self.json = []

    def convert_to_json(self):
        parser = ParseExcel(self.file_path, date_fields=['dob'])
        self.headers, self.data = parser.read_excel()
        self.construct_json()
        self.sanitize_json()
        json_data = json.dumps(self.json)
        return json_data

    def construct_json(self):
        for data in self.data:
            json = {}
            for i in range(0, len(data)):
                json.update({self.headers[i] : data[i]})
            self.json.append(json)

    def sanitize_json(self):
        for data in self.json:
            for key, value in data.items():
                if key in self.int_fields and value != '':
                    data[key] = int(value)
                elif key not in self.float_fields:
                    if isinstance(value, float):
                        data[key] = str(int(value))
