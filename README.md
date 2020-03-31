### ExceltoJSON 
#### This python package is a conversion program for excel to json.

### Usage
At present, this package has not been pushed to PyPI. So please download or clone this repository for using
in your project.

Example:

```
from exceltojson import ExcelToJson

convertor = ExcelToJson(file_path)
json_data = convertor.convert_to_json()

```
This will give output as a JSON string.

#### Optional Parameters
```ExceltoJson(file_path, index, name, int_fields, float_fields, date_fields)```

- file_path : Path to the excel file
- index : Excel sheet in the workbook. Give 0 for first sheet
- name : Name of the sheet
- int_fields : List of names of fields which are to be considered as integers
- float_fields : List of names of fields which are to be considered as floats
- date_fields : List of names of date fields. Have to be defined for date columns. Otherwise, they would be 
converted to floats
