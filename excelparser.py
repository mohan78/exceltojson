import xlrd


class ParseExcel:

    def __init__(self, source_path, sheet_index=0, sheet_name=None, date_fields=[]):
        self.source_path = source_path
        self.sheet_index = sheet_index
        self.sheet_name = sheet_name
        self.date_fields = date_fields
        self.sheet = None
        self.headers = []
        self.data = []

    def read_excel(self):
        workbook = xlrd.open_workbook(self.source_path)
        workbook_datemode = workbook.datemode
        if self.sheet_name:
            self.sheet = workbook.sheet_by_name(self.sheet_name)
        else:
            self.sheet = workbook.sheet_by_index(self.sheet_index)
        for col in range(0, self.sheet.ncols):
            self.headers.append(self.sheet.cell(0, col).value)

        date_field_indices = []
        for field in self.date_fields:
            if field in self.headers:
                date_field_indices.append(self.headers.index(field))

        for row in range(1, self.sheet.nrows):
            temp = []
            for col in range(0, self.sheet.ncols):
                value = self.sheet.cell(row, col).value
                if col in date_field_indices:
                    value = self.get_date_from_float(value, workbook_datemode)
                    temp.append(value)
                else:
                    temp.append(value)
            self.data.append(temp)
        return self.headers, self.data

    def get_date_from_float(self, value, datemode):
        y, m, d, h, mi, s = xlrd.xldate_as_tuple(value, datemode)
        if int(h)+int(mi)+int(s) == 0:
            return f"{y}-{m}-{d}"
        else:
            return f"{y}-{m}-{d} {h}:{mi}:{s}"