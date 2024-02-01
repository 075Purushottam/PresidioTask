from django.shortcuts import render
import pandas as pd
from django.shortcuts import redirect
import os


def home(request):
    project_root = os.path.dirname(os.path.realpath(__file__))

    excel_file_path = os.path.join(project_root, 'Employees.xlsx')
    df = pd.read_excel(excel_file_path)
    employees = df.to_dict(orient='records')
    row_count = len(df)
    context = {
        'employees':employees,
        'row_count':row_count
    }
    search_query = request.GET.get('search', '')
    age_query = request.GET.get('age', '')
    dob_query = request.GET.get('dob', '')
    salary_query = request.GET.get('salary', '')
    department_query = request.GET.get('department', '')

    if request.method == 'GET' and any([search_query, age_query, dob_query, salary_query, department_query]):
        df = filter_data(df, search_query, age_query, dob_query, salary_query, department_query)
        employees = df.to_dict(orient='records')
        context['employees'] = employees

    if request.method == 'GET':
        if 'total_average_salary' in request.GET:
                total_average_salary = df['Salary'].mean()
                context['total_average_salary'] = total_average_salary
                

    if request.method == 'POST':
        if 'submit_average_salary' in request.POST:
            department = request.POST.get('department')
            department_df = df[df['Department'] == department]

            if not department_df.empty:
                average_salary = department_df['Salary'].mean()

                context['department_average_salary'] = {
                    'department': department,
                    'average_salary': average_salary
                }

    
    return render(request,'mainapp/homepage.html',context)


def filter_data(df, search_query, age_query, dob_query, salary_query, department_query):
    filtered_df = df.copy()

    if search_query:
        filtered_df = filtered_df[filtered_df['FullName'].str.contains(search_query, case=False)]

    if age_query:
        filtered_df = filtered_df[filtered_df['Age'] == int(age_query)]

    if dob_query:
        filtered_df = filtered_df[pd.to_datetime(filtered_df['DOB']) == pd.to_datetime(dob_query)]

    if salary_query:
        filtered_df = filtered_df[filtered_df['Salary'] == int(salary_query)]

    if department_query:
        filtered_df = filtered_df[filtered_df['Department'].str.contains(department_query, case=False)]

    return filtered_df

def update_employee(request, employee_id):
    project_root = os.path.dirname(os.path.realpath(__file__))
    excel_file_path = os.path.join(project_root, 'Employees.xlsx')
    df = pd.read_excel(excel_file_path)
    employees = df.to_dict(orient='records')
    employee = next((emp for emp in employees if emp['id'] == int(employee_id)), None)
    index_to_update = df[df['id'] == employee_id].index

    if request.method == 'POST':
        full_name = request.POST.get('full_name')
        age = int(request.POST.get('age'))
        dob = request.POST.get('dob')
        salary = int(request.POST.get('salary'))
        department = request.POST.get('department')

        df.loc[index_to_update, ['FullName', 'Age', 'DOB', 'Salary', 'Department']] = [full_name, age, dob, salary, department]

        df.to_excel(excel_file_path, index=False)

        return redirect('home')
    
    return render(request, 'mainapp/update_employee.html',{'employee':employee})

def delete_employee(request, employee_id):
    project_root = os.path.dirname(os.path.realpath(__file__))
    excel_file_path = os.path.join(project_root, 'Employees.xlsx')
    df = pd.read_excel(excel_file_path)
    index_to_delete = df[df['id'] == employee_id].index
    df.drop(index_to_delete, inplace=True)
    df.to_excel(excel_file_path, index=False)
    employees = df.to_dict(orient='records')
    return render(request, 'mainapp/homepage.html',{'employees': employees})



