from django.shortcuts import render
import pandas as pd
import os

def home(request):
    # excel_file_path = '..\EmployeeManagement\Employees.xlsx'
    project_root = os.path.dirname(os.path.realpath(__file__))

    excel_file_path = os.path.join(project_root, 'Employees.xlsx')
    df = pd.read_excel(excel_file_path)
    
    employees = df.to_dict(orient='records')

    
    return render(request,'mainapp/homepage.html',{'employees': employees})
