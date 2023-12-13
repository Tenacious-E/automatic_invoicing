import pandas as pd 
import os
from os.path import exists
import datetime

"""
This program takes a spreadsheet with a company's clients rate per hour, hourly minimums, supervisor rates and 
creates multiple invoices with line items for each invoice.
"""


#need to have subfolders for client rates, output, and timesheet report.

#gets the client rates from client rate folder
Input_client_rates = os.path.normpath(os.getcwd() + os.sep + os.pardir + "/automatic_invoicing/Client Rates Folder")


#gets the timesheet report from timesheet folder
Input_report_timesheet = os.path.normpath(os.getcwd() + os.sep + os.pardir + "/automatic_invoicing/Timesheet Report")

#saves the output path
Output = os.path.normpath(os.getcwd() + os.sep + os.pardir + "/automatic_invoicing/Output")

#get the path in order to see if there's an error later
Error_file_path = os.path.normpath(os.getcwd() + os.sep + os.pardir)

#check if the path has an error
if exists(Error_file_path +'/rename_companies.csv'):
    os.remove(Error_file_path+'/rename_companies.csv')
#saving paths of client rates and timesheet
#filelist_client_rates = os.listdir(Input_client_rates)
#filelist_report_timesheet = os.listdir(Input_report_timesheet)

# create a dataframe for timesheet 
df= pd.read_excel(Input_report_timesheet +"/timesheet_report.xlsx")

#input client rates spreadsheet as a dataframe
rates_df= pd.read_excel(Input_client_rates +"/Current Client Rates - Client Minimum - Example Data.XLSX")

# now get the columns I want to work with
timesheets_df = df.loc[:len(df) -2 ,['employee','date','location','start time','end time','shift_title']]

#find unique names of companies in the rates spreadsheet and strip spaces so I can identify spelling errors later
companies_in_rates = rates_df['Location'].dropna().unique()
companies_in_rates = [i.lower().rstrip(' ').lstrip(' ') for i in companies_in_rates]

#create dictionaries so that each company can have a corresponding rate, hourly min, armed guard rate, and supervisor rate
company_rates_dict = dict()
hourly_min_dict = dict()
armed_guard_dict = dict()
supervisor_dict = dict()
for company,rate,hourly_min,armed_guard_rate,supervisor_rate in zip(rates_df['Location'],rates_df['Hourly Rate'],rates_df['Hourly Min'],rates_df['Supervisor Rate'],rates_df['Supervisor Rate']):
    comp  = company.lower().rstrip(' ').lstrip(' ')
    company_rates_dict[comp] = float(rate)
    hourly_min_dict[comp] = float(hourly_min)
    armed_guard_dict[comp] = float(armed_guard_rate)
    supervisor_dict[comp] = float(supervisor_rate)
#find the companies names in the timesheets spreadsheet and get unique company names and strip spaces in order to find spelling errors later
companies_in_timesheets = df['location'].dropna().unique()
companies_in_timesheets = [i.lower().rstrip(' ').lstrip(' ') for i in companies_in_timesheets]

#make copies of the dataframes
companies_in_timesheet_copy = companies_in_timesheets.copy()
companies_in_rates_copy = companies_in_rates.copy()

#function that takes a time as a string and converts it into a datetime and then returns a list of datetimes to be used in a dataframe later in the program
def change_time_to_correct_format(df,col_name):
    temp = pd.DataFrame(index = range(0,len(df)))
    #need to fill in 0s temporarily
    temp[col_name] = 0
    for i,line in enumerate(df[col_name]):
        temp.iloc[i,0]= datetime.datetime.strptime(line, "%I:%M %p")
    return temp[col_name].tolist()

#function finds the total hours and minutues from a string date. It then returns the total hours and minutes in the format: hours:minutes
def datetime_to_str(df,col_name):

    temp_df = pd.DataFrame(index = range(0,len(df)))
    temp_df[col_name] = 0
    for i,time in enumerate(df[col_name]):
        number_of_days = abs(int(str(time)[0:2]))
        if number_of_days >=  2:
            additional_hours =(number_of_days - 1) * 24
            hours = int(str(time)[-8:-6])
            hours += additional_hours
            time = str(time)
            time[-8:-6] = str(hours)
            total_hours = str(time)[-8:-6]
            total_minutes = str(time)[-5:-3]
            temp_df.iloc[i,0] = total_hours+':'+total_minutes
        else:
            total_hours = str(time)[-8:-6]
            total_minutes = str(time)[-5:-3]
            temp_df.iloc[i,0] = total_hours+':'+total_minutes
    return temp_df

#takes a string of times and converts them hourly times as ints              
def string_to_int_for_time_dif(df):
    temp_df = pd.DataFrame(index = range(0,len(df)))
    #temporarily use 0s
    temp_df['int_time'] = 0
    for i, time in enumerate(df):
        temp_df.iloc[i,0] = int(time[0:2]) + int(time[3:])/60
    return temp_df

#looking for spelling errors on the client rates spreadsheet
for c1 in companies_in_rates:
    for c2 in companies_in_timesheets:
        if c1.lower().strip() == c2.lower().strip():
            companies_in_rates_copy.remove(c1)
            companies_in_timesheet_copy.remove(c2)
#if there is a company that has a spelling error return a spreadsheet that states: need to add or update name in clients rate file and then ends the program
if len(companies_in_timesheet_copy) > 0:
    spell_check_df = pd.DataFrame([companies_in_timesheet_copy],index=["add or update name in the client rates file."])
    spell_check_df.to_csv(Error_file_path +'/rename_companies.csv')
#if there are no spelling errors program continues
else:

    # now iterate through each compaines billable items and create an invoice
    for company in timesheets_df['location'].unique():
        masked_company_df = timesheets_df[timesheets_df['location'] == company]
        # create a copy of the company dataframe
        temp_company_df = masked_company_df.copy()
        #call functions to find start and end times so they are in a usable format
        start_time_col = change_time_to_correct_format(temp_company_df,'start time')
        end_time_col = change_time_to_correct_format(temp_company_df,'end time')
        temp_company_df['start time'] = start_time_col
        temp_company_df['end time'] = end_time_col
        #create a column with net time difference between the end and start time of the employee. Needed it to be in datetime so it can be computed
        temp_company_df['net time difference'] = temp_company_df['end time'].values- temp_company_df['start time'].values

        #calls function that moves the net time difference column back to a string so it can be displayed in a nice manner later
        net_time_diff_in_str = datetime_to_str(temp_company_df,'net time difference')
        s = net_time_diff_in_str.copy()
        temp_company_df['net time difference'] = s.values

        # converts columns back to strings
        temp_company_df['start time'] = temp_company_df['start time'].astype(str)
        temp_company_df['end time'] = temp_company_df['end time'].astype(str)
        lst_start_time_str = []
        lst_end_time_str = []
        #get the start and endtimes from the strings so it can be displayed in a nice format
        for start, end in zip(temp_company_df['start time'],temp_company_df['end time']):
            lst_start_time_str.append(start[-8:])
            lst_end_time_str.append(end[-8:]) 
        
        temp_company_df['start time'] = lst_start_time_str
        temp_company_df['end time'] = lst_end_time_str

        #now convert the net time in hours from a string to an int so it can be displayed nicely in the output dataframe
        int_time_diff_col = string_to_int_for_time_dif(temp_company_df['net time difference'])
        temp_company_df['net time in hours'] = int_time_diff_col.values
        
        #this checking to see if the billed amount is less than the minimun allowed amount. If the billed amount is less than the minimum allowed, code changes billed amount to minimum allowed,
        temp_company_df_copy = temp_company_df.copy()
        for i, time in enumerate(temp_company_df_copy['net time in hours']):
            if temp_company_df.iloc[i,7] < hourly_min_dict[company.lower().rstrip(' ').lstrip(' ')]:
                temp_company_df.iloc[i,7] = hourly_min_dict[company.lower().rstrip(' ').lstrip(' ')]
       
        temp_company_df['rate'] = company_rates_dict[company.lower().rstrip(' ').lstrip(' ')]
        
        #finds the indices where there is an armed guard.
        indices_of_armed_guards = masked_company_df[masked_company_df['shift_title'].str.lower()=='armed'].index
        
        #charges the armed guard rate for armed guards
        for index in indices_of_armed_guards:
            temp_company_df.loc[index,'rate'] = armed_guard_dict[company.lower().rstrip(' ').lstrip(' ')]
        
        #finds the indices where there is a supervisor
        indices_of_supervisors =  masked_company_df[masked_company_df['shift_title'].str.lower().str.contains('supervisor') == True].index
        

        #charges the supervisor rate for the supervisors
        for index in indices_of_supervisors:
            temp_company_df.loc[index,'rate'] = supervisor_dict[company.lower().rstrip(' ').lstrip(' ')]
        
        #calculates the charges for each line item
        temp_company_df['charge'] = round(temp_company_df['net time in hours'] * temp_company_df['rate'], 2) 
        
        #drop index so it looks better on the output
        temp_company_df = temp_company_df.reset_index().drop(['index'],axis = 1)
        
        #creates a dataframe with the relevant columns to be displayed 
        df_final_sum = pd.DataFrame({'employee':'','date':'','location':'','start time':'','end time':'','net time difference':'','net time in hours':'Total','charge':round(sum(temp_company_df['charge']),2)},index=[0])

        #creates the final dataframe with the relevant column names and values.
        final_df = pd.concat([temp_company_df, df_final_sum], ignore_index = True, axis = 0)

        #change the date as the index
        final_df.index = final_df['date'].values
        #drop the column date because its in the index now
        final_df.drop(['date'], axis = 1, inplace =True)
        print(Output + "/" + company + ' invoice.xlsx')
        
        #send the dataframe to the correct file as an excel file.
        final_df.to_excel(Output + "/" + company + ' invoice.xlsx')
        

