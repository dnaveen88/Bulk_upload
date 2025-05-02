from django.http import HttpResponse
import os
import pandas as pd
from wsgiref.util import FileWrapper
from xlsxwriter.utility import xl_rowcol_to_cell
import datetime

def create_filepath(module, filename):
    current_path = os.getcwd()
    dt = str(datetime.datetime.now())
    file_path = current_path + '/'+module+'/excel_templates/'+filename+dt+'.xlsx'
    return file_path


def download_template(file_path):
    with open(file_path,'rb') as  file_obj:
        response=HttpResponse(FileWrapper(file_obj),content_type='application/xls')
        filename = os.path.split(file_path)[-1]
        response['Content-Disposition'] = "attachment; filename=%s"%filename
        response['Content-Length']  = os.path.getsize(file_path)
    return response


def data_validation( col_num, worksheet, params=None, dropdown_range=None, special_key=None):
    for i in range(502):
        try:
            if dropdown_range!=None:
                input_message = 'If yes,fill minimum and maximum duration' if special_key == 'Include Lateral Entry' else ''
                worksheet.data_validation(xl_rowcol_to_cell(i,col_num),{
                                        'validate':'list',
                                        'source':dropdown_range,
                                        'input_message': input_message,
                                        'criteria' : '=isblank()=False'
                            })  

            elif params[0] == 'integer' :
                input_message = ''
                if special_key == 'Exam Duration':
                    input_message  ='between 0 and 360'
                else:
                    msg = {'Maximum Duration':['Maximum duration', 'Minimum duration'], 'Maximum Marks':['Maximum marks', 'Minimum marks'], 'Maximum Credit':['Maximum credit', 'Minimum credit']}
                    for key , val in msg.items():
                        input_message = val[0] +' should be greater than '+ val[1] if special_key == key else 'between 0 and 200'
                worksheet.data_validation(xl_rowcol_to_cell(i,col_num), {'validate': params[0],
                                        'criteria': '>',
                                        'value': -1,
                                        'error_message': 'It should be an integer',
                                        'input_title': 'Enter an integer:',
                                        'input_message': input_message})

            elif params[0] == 'date':
                worksheet.data_validation(xl_rowcol_to_cell(i,col_num),{
                            'validate':'date',
                            'criteria': 'between',
                            'minimum': datetime.date(1990, 1, 1),
                            'maximum': datetime.date(2050, 12, 12),
                            'input_title': 'Format:',
                            'error_message': 'It should be DD/MM/YYYY format',
                            'input_message': 'DD/MM/YYYY'}) 

            elif params[0] == 'decimal':
                        worksheet.data_validation(xl_rowcol_to_cell(i,col_num),{
                            'validate':'decimal',
                            'criteria': '>',
                            'value': 0.0})  

            else:
                pass
        except:
            pass
    return('success')

def hidden_column_dropdown(worksheet, vals, row, sheet_name):
    b=0
    last_col = (len(vals))-1
    for i in vals:
        worksheet.write(b,int(row),i)
        b=b+1
    dropdown_range = "="+sheet_name+"!"+xl_rowcol_to_cell (0,int(row), row_abs=True, col_abs=True)+":"+xl_rowcol_to_cell (last_col,int(row), row_abs=True, col_abs=True)
    return dropdown_range

def merge(worksheet, initial_row, initial_col, final_row, final_col, j, header_fmt):
    worksheet.merge_range(initial_row, initial_col, final_row, final_col, j, header_fmt)


def data_attributes(column_name, define_data_type):
    params = []
    data_types = define_data_type()
    for entry in data_types:
        for key, val in entry.items():
            if column_name in  val:
                params.append(key)
    return params

def excel_formats():
    header_format = {
            'border':1,
            'text_wrap':True,
            'valign':'vcenter',
            'indent':1,
            'bold': True,
            'bg_color':'#D7E4BC',
            'locked': True
            }
    merge_format={'bold':  True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color':'#D7E4BC',
            'locked': True}
    mandatory_format={
            'border':1,
            'text_wrap':True,
            'valign':'vcenter',
            'indent':1,
            'bold': True,
            'bg_color':'#F08080',
            'locked': True
            }
    return [header_format, merge_format, mandatory_format]