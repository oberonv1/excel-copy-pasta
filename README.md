# excel-copy-pasta
Python script copy and paste from separate .xlsx using openpyxl library

To run script:
- create `input_templates` and `output_templates` folder
- then run the following commands
```sh
 pip install openpyxl
 python flatness_transpose.py -i 'input.xlsx' -o 'output.xlsx' -s 'input_sheet' -S 'output_sheet'
```

Currently only supports .xlsx
