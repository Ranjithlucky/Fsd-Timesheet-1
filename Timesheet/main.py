import pandas as pd
import datetime
from fastapi import FastAPI, UploadFile,File
import calendar
from openpyxl.utils import get_column_letter

app = FastAPI()

@app.post("/timesheet")
def generate_timesheet(file1: UploadFile = File(...),file2: UploadFile = File(...)):

    #input-1<<<<<<<<-------------------------File1 -------------------------------------------->>>>>>>>
    df = pd.read_excel(file1.file.read(),  keep_default_na=False, parse_dates=True)
    billable_data = df[df["Billing Action"] == "Billable"]
    billable_data["Date"] = pd.to_datetime(billable_data["Date"], format="%Y-%m-%d")
    billable_data['Time Quantity'] = billable_data['Time Quantity'].replace(10, 9)

    # get the year, month, date
    year = billable_data["Date"].dt.year.unique()[0]
    month = billable_data["Date"].dt.month.unique()[0]
    num_days = calendar.monthrange(year, month)[1]

    # create a column list format
    dates = []
    for day in range(1, num_days+1):
       date = datetime.datetime(year, month, day)
       if date.strftime("%b-%d") not in dates:
            dates.append(date.strftime("%b-%d"))

    
    
    # weekend for to give css
    weekend_dates = []
    for day in range(1, num_days + 1):
        if datetime.datetime(year, month, day).weekday() in [5, 6]:
            weekend_dates.append(datetime.datetime(year, month, day).strftime("%b-%d"))

    

   # input2<<<<<<<<-------------**** File2*****--------------------------------------------->>>>>>>>
    df2 = pd.read_excel(file2.file.read(), keep_default_na=False, parse_dates=True)
    location_data = df2[df2["Off/On"] == "Offshore"]
    location_data2 = df2[df2["Off/On"] == "Onsite"]
    location_data3 = pd.concat([location_data, location_data2])

    #  # input3<<<<<<<<-------------**** File2*****--------------------------------------------->>>>>>>>
    # df3 = pd.read_excel(file3.file.read(), keep_default_na=False, parse_dates=True)
    # project=df3[df3["Project Id"]==1000212211]
    

    
    #merged both file1 and file2
    merged_data = pd.merge(billable_data, location_data3,on="Employee ID", how="left")
    # merged_data = pd.merge(pd.merge(billable_data, location_data3, on="Employee ID"), project, on="Project Id", how="left")

    # unique values of both 1 &2
    unique_employee_ids = merged_data["Employee ID"].unique()
    
    # Create a new dataframe 
    result = pd.DataFrame(columns=["Sl No",
                                  "Project",
                                   "Project Name",
                                   "Employee ID", 
                                   "Name", 
                                   "ON/OF", 
                                   "Location",
                                   "SOW",
                                   "PO",
                                   "Total_Worked_Days",
                                   "Total_billable_hours",
                                    "Hourly_rate",
                                    "Total cost"] + dates)
    
    
    
    # Loop through each unique "Employee ID"
    for i, employee_id in enumerate(unique_employee_ids):
        employee_data = merged_data[merged_data["Employee ID"] == employee_id]
       
        #total working days
        total_worked_days = 0
        for hours in employee_data["Time Quantity"]:
            if hours == 8:
               total_worked_days += 1
            elif hours == 4.5:
              total_worked_days += 0.5
            elif hours == 9:
              total_worked_days += 1
            else:
              total_worked_days += 0.5
            
        
         # content in the row
        first_row = employee_data.iloc[0]
        
        #location as time quantity
        if first_row["Location"] == "Gurgaon":
            employee_data["Time Quantity"] = 8
        elif first_row["Location"] == "Kolkata":
            employee_data["Time Quantity"] = 8
        elif first_row["Location"] == "Noida":
            employee_data["Time Quantity"] = 8
        elif first_row["Location"] == "Onsite":
            employee_data["Time Quantity"] = 8
        
        #total hours
        total_hours=employee_data["Time Quantity"].sum()
        

        #input 2 <<<<<<<<<<<<<<<<<---------------------------------->>>>>>>>>>>>>>>>>>>>>>>>>
        # ------hourly rate---******************************---------
        hourly_rate = (first_row["Rate"])
        hourly_rate_formatted = ("${:.2f}".format(hourly_rate))
        hourly=hourly_rate_formatted.replace('$nan', '0')


        #---------------Total_cost--------------*********************-------------------
        total_cost = float(total_hours) * hourly_rate
        cost = ("${:.2f}".format(total_cost))
        costly=cost.replace('$nan', '0')
        

       
        #genereate sequence of both 1&2
        result_row = {"Sl No": i+1,
                      "Project": first_row["Project"],
                      "Project Name": "AGERO FULL STACK DEV T&M",
                      "Employee ID": first_row["Employee ID"],
                      "Name": first_row["Name"],
                      "ON/OF": first_row["ON / OF"],
                      "Location":first_row["Location"],
                      "SOW": first_row["SOW"],
                      "PO": first_row["PO"],
                      "Total_Worked_Days":total_worked_days,
                      "Total_billable_hours":total_hours,
                      "Hourly_rate":hourly,
                      "Total cost":costly
                      }
        
        # Add the result row to the result dataframe
        result = result.append(result_row, ignore_index=True)

       

        # date fetch time quantity for sum method as billable
        for date in dates:
            grouped_data = employee_data.groupby(["Date", "Employee ID"])["Time Quantity"].sum().reset_index()
            date_data = grouped_data[grouped_data["Date"].dt.strftime("%b-%d") == date]
            if not date_data.empty:
                time_quantity = date_data["Time Quantity"].iloc[0]
                result.loc[result["Employee ID"] == employee_id, date] = time_quantity
            else:
                result_row[date] = " "
    
    # CSS -START
    # Function to apply background color to cells
    def color_background(value):
       if value==9:
        color = 'None'
       elif value == 4.5:
          color = 'yellow'
       elif value==8:
          color='None'
       else:
          color = 'red'
       return f'background-color: {color}'   
    styler = result.style.set_properties(**{'text-align': 'center'}).applymap(color_background, subset=pd.IndexSlice[:, dates])
    
    for date in weekend_dates:
        styler = styler.set_properties(**{'background-color': 'None'}, subset=date)
    
    #  result dataframe to a new excel file
    writer = pd.ExcelWriter("time.xlsx", engine='openpyxl')
    styler.to_excel(writer, sheet_name='Sheet1', index=False)

    # Access the worksheet and set column widths
    worksheet = writer.sheets['Sheet1']
    column_widths = [7, 15, 27, 12, 35, 8, 17, 18, 20, 20,20] 
    for i, width in enumerate(column_widths):
        column_letter = get_column_letter(i+1) 
        worksheet.column_dimensions[column_letter].width = width 

    writer.save()
    return "Timesheet has been created successfully"
    

