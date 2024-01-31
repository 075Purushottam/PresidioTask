from django.shortcuts import render
import pandas as pd

def home(request):
    excel_file_path = '..\EmployeeManagement\Employees.xlsx'

    df = pd.read_excel(excel_file_path)
    
    employees = df.to_dict(orient='records')

    
    return render(request,'mainapp/homepage.html',{'employees': employees})
