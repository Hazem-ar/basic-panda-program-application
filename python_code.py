import pandas as pd

origninal_excel_file= pd.read_excel('Employee Sample Data.xlsx')
Dforexcel = pd.DataFrame(origninal_excel_file)
print(Dforexcel)
#cleaning data: empty cells, Data in wrong format, wrong data, and duplicates 
print(Dforexcel.info())
print(len(Dforexcel))
Dforexcel.fillna({"Exit Date": "still in company"}, inplace=True)
print(Dforexcel)
print(Dforexcel.duplicated())

cleaned_dataframe = Dforexcel.dropna()
print(cleaned_dataframe)
print(len(Dforexcel))
#data cleaned 

#change any values in the first 5 rows
cleaned_dataframe.loc[0, "Department"]= "Finance"
cleaned_dataframe.loc[1, "Business Unit"]= "Speciality Products"
cleaned_dataframe.loc[2, "Ethnicity"]= "Asian"
cleaned_dataframe.loc[3, "Age"]= 55.0
cleaned_dataframe.loc[4, "Hire Date"]= '2011-01-01 00:00:00'
print(cleaned_dataframe, "\n")

#max salary
max_salary= cleaned_dataframe["Annual Salary"].idxmax()
print(cleaned_dataframe.loc[[max_salary]], "\n")



grouped_by_dep= cleaned_dataframe.groupby("Department")
average_age= grouped_by_dep['Age'].mean()
average_salary= grouped_by_dep['Annual Salary'].mean()
print("average age by department: \n",average_age, "\n","average salary by department: \n",average_salary)

grouped_by_dep_and_ethnicity =cleaned_dataframe.groupby(["Department", "Ethnicity"])
max_age= grouped_by_dep_and_ethnicity['Age'].max()
min_age= grouped_by_dep_and_ethnicity["Age"].min()
median_salary= grouped_by_dep_and_ethnicity["Annual Salary"].median()
print("max age: \n", max_age, "min age: \n", min_age, "median salary: \n", median_salary)

file_name = "cleaned_employee_sample_data.xlsx"

cleaned_dataframe.to_excel(file_name)

