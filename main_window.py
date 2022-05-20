import PySimpleGUI as sg
import csv, os
import sqlite3
import pandas as pd

working_directory = os.getcwd()
data = []
# declaring column names for the pysimplegui table
header_list = ['Counties', 'Municipalities', 'Application', 'Project Title',
       'District Priority', 'Municipal Priority', 'Rating',
       'Type Of Improvement', 'Sectionalized', 'Total Requested Amount',
       'Total Estimated Cost', 'Eligible Amount', 'Urban Aid',
       'Recomended Total','Distributed Total']

#creates initial table from excel sheet 
def create_table(path):
    df = pd.read_excel(path,header=0)
    connection = sqlite3.connect('test.db')
    df.to_sql(
            name="Municipal Aid", con=connection,if_exists='replace',index=False,dtype={'Counties' : 'TEXT',
                'Municipalities': 'TEXT',
                'Application': 'TEXT',
                'Project Title': 'TEXT',
                'District Priority' : 'TEXT',
                'Municipal Priority	Rating' : 'INTEGER',
                'Type Of Improvement' : 'TEXT',
                'Sectionalized' : 'TEXT',
                'Total Requested Amount': 'INTEGER',
                'Total Estimated Cost': 'INTEGER',
                'Eligible Amount': 'INTEGER',
                'Urban Aid':'INTEGER',
                'Recommended_Total':'INTEGER'}
        )
    connection.commit()
    connection.close()
## overrides table with dataframe for exporting the excel sheet
def create_df_table(df):
    connection = sqlite3.connect('test.db')
    df.to_sql(
            name="Municipal Aid", con=connection,if_exists='replace',index=False,dtype={'Counties' : 'TEXT',
                'Municipalities': 'TEXT',
                'Application': 'TEXT',
                'Project Title': 'TEXT',
                'District Priority' : 'TEXT',
                'Municipal Priority	Rating' : 'INTEGER',
                'Type Of Improvement' : 'TEXT',
                'Sectionalized' : 'TEXT',
                'Total Requested Amount': 'INTEGER',
                'Total Estimated Cost': 'INTEGER',
                'Eligible Amount': 'INTEGER',
                'Urban Aid':'INTEGER',
                'Recommended Total':'INTEGER',
                'Distributed Total': 'INTEGER'}
        )
    connection.commit()
    connection.close()
#prints database to excel
def to_excel(path,name):
    connection = sqlite3.connect('test.db')
    df = pd.read_sql_query("SELECT * FROM 'Municipal Aid'",connection)
    df.to_excel(path + '/' +name+".xlsx")
    connection.commit()
    connection.close()

#creates table to display
def create_data_frame():
    connection = sqlite3.connect('test.db')
    data_frame = pd.read_sql_query("SELECT * FROM 'Municipal Aid'",connection)
    connection.commit()
    connection.close()
    return data_frame

## get number of points from database
def get_points():
    connection = sqlite3.connect('test.db')
    cursor = connection.cursor()
    query = "SELECT SUM(Rating) FROM 'Municipal Aid' WHERE \"Municipal Priority\"=1"
    sum = cursor.execute(query)
    sumTest = sum.fetchone()
    connection.close()
    return sumTest[0]
# creates virtual column for total requested amount (1 point allotment * rating)
def create_recommended_total(int):
    int = int
    connection = sqlite3.connect('test.db')
    #sqlite limitation means that only a virtual column can be added with ALTER TABLE
    connection.execute(f"ALTER TABLE 'Municipal Aid' ADD COLUMN \"Recomended Total\" Integer GENERATED ALWAYS AS ({int}*Rating);")
    connection.commit()
    connection.close()
##function to create a new column that will be added to the pandas data frame based on ratings input ranges
def point_ranges(up1,up2,up3,up4,lp1,lp2,lp3,lp4,amt1,amt2,amt3,amt4,values):
    new_column = []
    for i in values:
        if i[6] <= int(up1) and i[6] >= int(lp1):
            new_column.append(i[13]+int(amt1))

        elif i[6] <= int(up2) and i[6] >= int(lp2):
            new_column.append(i[13]+int(amt2))

        elif i[6] <= int(up3) and i[6] >= int(lp3):
            new_column.append(i[13]+int(amt3))

        elif i[6] <= int(up4) and i[6] >= int(lp4):
            new_column.append(i[13]+int(amt4))
        else:
            new_column.append(0)

    return new_column


sg.theme('DefaultNoMoreNagging')

#window
layout = [
    ##this is the data structure that determines how the UI looks, each [] is a line, the keys are referenced in the while loop below to do actions to/from elements
    [sg.Text("Choose a xlsx file:")],
    [sg.InputText(key="-FILE_PATH-"), sg.FileBrowse(initial_folder=working_directory, file_types = [("xlsx Files","*.xlsx")])],
    [sg.Text("Number of #1 points:"),sg.Text(size=(5, 1),key='-ONE-POINTS-'),sg.Text("Total Allotment"),sg.Input(key='-TA-', do_not_clear=True,size=(15,1))],
    [sg.Text("#1 Point Value:"),sg.Text(size=(7,1),key='-POINT-VALUE-')],
    [sg.Text("Upper Point Range"),sg.Input(key='-up1-',size=(4,1)),sg.Text('Lower Point Range'),sg.Input(key='-lp1-',size=(4,1)),sg.Text('Amount'),sg.Input(key="-amt1-",size=(8,1))],
    [sg.Text("Upper Point Range"),sg.Input(key='-up2-',size=(4,1)),sg.Text('Lower Point Range'),sg.Input(key='-lp2-',size=(4,1)),sg.Text('Amount'),sg.Input(key="-amt2-",size=(8,1))],
    [sg.Text("Upper Point Range"),sg.Input(key='-up3-',size=(4,1)),sg.Text('Lower Point Range'),sg.Input(key='-lp3-',size=(4,1)),sg.Text('Amount'),sg.Input(key="-amt3-",size=(8,1))],
    [sg.Text("Upper Point Range"),sg.Input(key='-up4-',size=(4,1)),sg.Text('Lower Point Range'),sg.Input(key='-lp4-',size=(4,1)),sg.Text('Amount'),sg.Input(key="-amt4-",size=(8,1))],
    [sg.Table(values=data,headings=header_list,display_row_numbers=True,
                  auto_size_columns=True,vertical_scroll_only = False,num_rows=35,key="-table1-",visible=False)],
    [sg.Text("Sum of Recommended Total: "),sg.Text(size=(9, 1),key='-RT-'),sg.Text("Sum of Distributed Total: "),sg.Text(size=(9, 1),key='-DT-')],
    [sg.Text("Choose an export location:")],
    [sg.InputText(key="-EXPORT_PATH-"), sg.FolderBrowse(initial_folder=working_directory)],
    [sg.Text("Input filename:"),sg.Input(key="-file-name-",size=(20,1))],
    [sg.Button("Calculate"),sg.Button("Export"), sg.Exit()],
    [sg.Text("Status:"),sg.Text(size=(35,1),key='-STATUS-')]
]

window = sg.Window("Municipal Aid Point Calculator", layout,resizable=True)
#while loop for submitting the document
while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Exit'):
        break
## this is kinda long and ugly but its basically running the program based off the calcculate button
    elif event == "Calculate":
        try:
            csv_address = values["-FILE_PATH-"]
            create_table(csv_address)
            one_points = get_points()
            window['-ONE-POINTS-'].Update(one_points)
            ta = values["-TA-"]
            ##strips commas from answer
            total_allotment = ta.replace(",","")
            point_value = round((int(total_allotment)/one_points),0)
            window['-POINT-VALUE-'].update(point_value)
            ##working on table here, populating data frame with table from sql
            create_recommended_total(point_value)
            dataframe = create_data_frame()
            data_frame_list = dataframe.values.tolist() 
            #creating new column with distributed point ranges
            new_column = point_ranges(values['-up1-'],values['-up2-'],values['-up3-'],values['-up4-'],
            values['-lp1-'],values['-lp2-'],values['-lp3-'],values['-lp4-'],
            values['-amt1-'],values['-amt2-'],values['-amt3-'],values['-amt4-'],data_frame_list)
            ##adding distributed total column to dataframe
            dataframe = dataframe.assign(Distributed_Total=new_column)      
            window["-table1-"].update(values=dataframe.values.tolist(),visible=True)
            window["-RT-"].update(str(dataframe['Recomended Total'].sum()))
            window["-DT-"].update(str(dataframe['Distributed_Total'].sum()))
            window['-STATUS-'].update("Calculated, press export for xlsx file")

        except:   
            window['-STATUS-'].update("Must Enter Values")
## elif for export button being pressed, creates a new dataframe with the two new columns and overrides the database so it can print it to excel
    elif event == "Export":
        try:   
            exp_path = values["-EXPORT_PATH-"]
            create_df_table(dataframe)
            try:
                to_excel(exp_path,values["-file-name-"])
                filename =values["-file-name-"]
                window['-STATUS-'].update(f"File exported as {filename} ")
            except:
                window['-STATUS-'].update("Must choose export folder/file name")
        except:
            window['-STATUS-'].update("Must calculate first")

        
window.close()





