# -*- coding: utf-8 -*-
"""
Created on Fri Jan 31 11:21:42 2020

@author: jwenger
"""

from __future__ import print_function

import tkinter as tk
import tkinter.filedialog as filedialog

import pandas as pd
import os
import matplotlib.pyplot as plt

from mailmerge import MailMerge

import comtypes.client

import fitz

reps = ''
program_data = ''
mail_merge_template = ''
source = ''
template = ''
    
#%%
def make_pdf(name, name2):
    word_input = name
    pdf_output = name2
    wdFormatPDF = 17
    
    in_file = word_input
    out_file = pdf_output
    
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()
    

def add_image(pdf_file, image):
    
    input_file = pdf_file
    output_file = pdf_file.strip(".pdf") + "with graph.pdf"
    image_file = image
    
    image_rectangle = fitz.Rect(250, 450, 350, 550)
    
    file_handle = fitz.open(input_file)
    first_page = file_handle[0]

    # add the image
    first_page.insertImage(image_rectangle, image_file)
    
    file_handle.save(output_file)
    

#%%
def make_word_docs(mail_merge_template_final):
    #os.chdir('C:\\Users\\jwenger\\Desktop\\automated_reports\\phase_2')
    
    sheet = mail_merge_template_final
    sheet = sheet.fillna('')
    
    for index, row in sheet.iterrows():      
        #template = r"C:\Users\jwenger\Desktop\automated_reports\phase_2\template docx includes picture.docx"
        
        document = MailMerge(template)
        print(document.get_merge_fields())                        
                        
        document.merge(CHAMBER = str(row['CHAMBER']),
                   Chamber1 = str(row['Chamber']),
                   District = str(row['District']),
                   ISBEIECAMYEAR = str(row['ISBEIECAMYEAR']),
                   ISBELegDataYear = str(row['ISBELegDataYear']),
                   NotServedByIsbe = str(row['NotServedByIsbe']),
                   PFAEserved = str(row['PFAEServed']),
                   PFAServed = str(row['PFAServed']),
                   PIServed = str(row['PIServed']),
                   PercentServedByISBERounded = str(row['PercentServedByISBERounded']),
                   PercentUnder5atFPL = str(row['PercentUnder5atFPL']),
                   PopYear = str(row['PopYear']),
                   ServedByISBE = str(row['ServedByISBE']),
                   Under5atFPL = str(row['Under5atFPL']),
                   M_1 = str(row['1']),
                   M_2 = str(row['2']),
                   M_3 = str(row['3']),
                   M_4 = str(row['4']),
                   M_5 = str(row['5']),
                   M_6 = str(row['6']),
                   M_7 = str(row['7']),
                   M_8 = str(row['8']),
                   M_9 = str(row['9']),
                   M_10 = str(row['10']),
                   M_11 = str(row['11']),
                   M_12 = str(row['12']),
                   M_13 = str(row['13']),
                   M_14 = str(row['14']),
                   M_15 = str(row['15']),
                   M_16 = str(row['16']),
                   M_17 = str(row['17']),
                   M_18 = str(row['18']),
                   M_19 = str(row['19']),
                   M_20 = str(row['20']),
                   M_21 = str(row['21']),
                   M_22 = str(row['22']),
                   M_23 = str(row['23']),
                   M_24 = str(row['24']),
                   M_25 = str(row['25']),
                   M_26 = str(row['26']),
                   M_27 = str(row['27']),
                   M_28 = str(row['28']),
                   M_29 = str(row['29']),
                   M_30 = str(row['30']),
                   M_31 = str(row['31']),
                   M_32 = str(row['32']),
                   M_33 = str(row['33']),
                   M_34 = str(row['34']),
                   M_35 = str(row['35']),
                   M_36 = str(row['36']),
                   M_37 = str(row['37']),
                   M_38 = str(row['38']),
                   M_39 = str(row['39']),
                   M_40 = str(row['40']),
                   M_41 = str(row['41']),
                   M_42 = str(row['42']),
                   M_43 = str(row['43']),
                   M_44 = str(row['44']),
                   M_45 = str(row['45']),
                   M_46 = str(row['46']),
                   M_47 = str(row['47']),
                   M_48 = str(row['48']),
                   M_49 = str(row['49']),
                   M_50 = str(row['50']),
                   M_51 = str(row['51']),
                   M_52 = str(row['52']),
                   M_53 = str(row['53']),
                   M_54 = str(row['54']),
                   M_55 = str(row['55']),
                   M_56 = str(row['56']),
                   M_57 = str(row['57']),
                   M_58 = str(row['58']),
                   M_59 = str(row['59']),
                   M_60 = str(row['60']),  
                   M_61 = str(row['61']),
                   M_62 = str(row['62']),
                   M_63 = str(row['63']),
                   M_64 = str(row['64']),
                   M_65 = str(row['65']),
                   M_66 = str(row['66']),
                   M_67 = str(row['67']),
                   M_68 = str(row['68']),
                   M_69 = str(row['69']),
                   M_70 = str(row['70']),
                   M_71 = str(row['71']),
                   M_72 = str(row['72']),
                   M_73 = str(row['73']),
                   M_74 = str(row['74']),
                   M_75 = str(row['75']),
                   M_76 = str(row['76']),
                   M_77 = str(row['77']),
                   M_78 = str(row['78']),
                   M_79 = str(row['79']),
                   M_80 = str(row['80']),
                   M_81 = str(row['81']),
                   M_82 = str(row['82']),
                   M_83 = str(row['83']),
                   M_84 = str(row['84']),
                   M_85 = str(row['85']),
                   M_86 = str(row['86']),
                   M_87 = str(row['87']),
                   M_88 = str(row['88']),
                   M_89 = str(row['89']),
                   M_90 = str(row['90']),
                   M_91 = str(row['91']),
                   M_92 = str(row['92']),
                   M_93 = str(row['93']),
                   M_94 = str(row['94']),
                   M_95 = str(row['95']),
                   M_96 = str(row['96']),
                   M_97 = str(row['97']),
                   M_98 = str(row['98']),
                   M_99 = str(row['99']),
                   image = str(row['image'])    
                   
                           )   
        word_doc_file = source  + r'/' + row['Chamber'] + str(row['District']) + '.docx'
        word_doc_temp = source  +  r'/' + row['Chamber'] + str(row['District']) + '.docx'
        pdf_file = source  + r'/' + row['Chamber'] + str(row['District']) + '.pdf'
        
        document.write(word_doc_file)
        document.write(word_doc_temp)
        make_pdf(word_doc_temp, pdf_file)
        add_image(pdf_file, row['image'])   




#%%
#%%
#%% methods:

def make_list_of_programs(subset_df):
        grabber = []
        for index, rows in subset_df.iterrows():
            grabber.append(rows.GrantFund)
            grabber.append(rows.OrganizationName)
            grabber.append(rows.GranteeCity)
        return(grabber)

#method to be applied to the names in the senator list. It should capitalize 
#all the names, remove the punctuation andsuffixes, and also change the first
# name into an initial    
def fix_names(name):
    name = name.upper() #capitalizes
    print(name)
    return(name)

#takes all the senators and fixes them, assigns new name to a new list 
def standardize_names(df):
    df_first_name = list(map(fix_names, df['First Name'])) 
    df_last_name = list(map(fix_names, df['Last Name'])) 
    
    df_initial = list(map(lambda first, last: first[0] + ' ' + last, 
                                df_first_name,
                                df_last_name))
    
    #takes the fixed and lists and subs thems into the data frame 
    
    df['First Name'] = df_first_name
    df['Last Name'] = df_last_name 
    df['Name'] = df_initial
    return(df)
    
def fix_program_data(program_data):
    program_data['GranteeCity'] = list(map(lambda x: x.title(), program_data['GranteeCity']))
    program_data['StateRepresentative'] =list(map(lambda x: x.upper(), program_data['StateRepresentative']))
    return(program_data)

# this functions tests to see if the name passed through is the same last name and same first intial as on of the existing senators from our database, 
# if both of these conditions are met, this functions returns the district associated with that senator
def match_district(name):
    for index, row in reps.iterrows():
        if row['Last Name'] in name and row['First Name'][0] == name[0]:
           return(row['District'])
    
# the next is to take the program information from the program data data frame and populate it into the mail_merge_template based off of the 

def fill_out_mail_merge_template(program_data, mail_merge_template):
    #step 1: create a list of all the districts needed
    districts = program_data['rep district'].unique()
    districts.sort()
    #step 1b. set the index of mailmerge df to the district names
    mail_merge_template = mail_merge_template.set_index('District', drop = False) 
    
    # step 2a. subset the program data df based on districts 
    for i in districts:
        subset_df = program_data[program_data['rep district'] == i] 
        #add all the program information for the district into a list called grabber
        grabber = make_list_of_programs(subset_df)
        counter = 1
        for info in grabber:
            mail_merge_template.at[i, str(counter)] = info
            counter += 1
    
    return mail_merge_template



def create_pie_charts(df):
    for index, row in df.iterrows():
        filename = row['image'] 
        percent_served = float(row['PercentServedbyISBE'].strip('%')) #takes the percent served by ISBE, removes the % sign and turns it into a float
        values = [percent_served, 100 - percent_served]
        labels = ['','']
        colors = ['lightgray', 'darkslateblue']

        fig1, ax1 = plt.subplots()
        fig1.set_size_inches(2,2)
                        
        ax1.pie(values,
                labels=labels,
                colors=colors,
                autopct='%1.1f%%',
                shadow=True, 
                startangle=90)
        
        ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
        fig1.savefig(filename, bbox_inches = 'tight')
        plt.close(fig1)
        
  
        
    
def make_images(df):
    df['image'] = source + "/chart" + (df['Chamber']) + df['District'].astype(str) +'.jpg'

    df['image'] = df['image'].str.replace('\\', '\\\\')

    print('boop')
    create_pie_charts(df)
    return(df)
    
    
#%%


    

#STANDARDIZING NAMES
def automate_data(reps, program_data, mail_merge_template, source):
    
    reps = standardize_names(reps)
    program_data = fix_program_data(program_data)
    program_data['rep district'] = list(map(match_district, program_data['StateRepresentative']))

    mail_merge_template = fill_out_mail_merge_template(program_data, mail_merge_template)

    mail_merge_template = make_images(mail_merge_template)
    mail_merge_template.to_excel('Mail Merge House Template.xlsx')
    
    location = source + "/Mail Merge SENATE Template.xlsx"
    
    mail_merge_template.to_excel(location)
    make_word_docs(mail_merge_template)


master = tk.Tk()
master.title('Ounce Data Sheets')
master.geometry('1000x700')

def input1():
    input_path = tk.filedialog.askopenfilename()
    input_entry1.delete(1, tk.END)  # Remove current text in entry
    input_entry1.insert(0, input_path)  # Insert the 'path'
    

def input2():
    input_path = tk.filedialog.askopenfilename()
    input_entry2.delete(1, tk.END)  # Remove current text in entry
    input_entry2.insert(0, input_path)  # Insert the 'path'

def input3():
    input_path = tk.filedialog.askopenfilename()
    input_entry3.delete(1, tk.END)  # Remove current text in entry
    input_entry3.insert(0, input_path)  # Insert the 'path'
    
def input_word_doc():
    input_path = tk.filedialog.askopenfilename()
    input_entry_word_doc.delete(1, tk.END)  # Remove current text in entry
    input_entry_word_doc.insert(0, input_path)  # Insert the 'path'

def output(): #the command associated with the "browse" button
    path = tk.filedialog.askdirectory()
    output_entry.delete(1, tk.END)  # Remove current text in entry
    output_entry.insert(0, path)  # Insert the 'path'

def begin():
    global reps 
    reps = pd.read_excel(input_entry1.get())
    global program_data
    program_data = pd.read_excel(input_entry2.get())
    global mail_merge_template
    mail_merge_template = pd.read_csv(input_entry3.get())
    global source
    source = output_entry.get()
    global template 
    template = input_entry_word_doc.get()
    
    automate_data(reps, program_data, mail_merge_template, source)

    

#%%
left_frame = tk.Frame(master)
right_frame = tk.Frame(master)
bottom_frame = tk.Frame(master)
#line = tk.Frame(master, height=1, width=400, bg="grey80", relief='groove')

#%%
input_path1 = tk.Label(left_frame, text="Link to the Excel Sheet 1: \n The List of house reps:")
input_entry1 = tk.Entry(left_frame, text="", width=40)
browse1 = tk.Button(left_frame, text="Browse", command=input1)

input_path2 = tk.Label(left_frame, text="Link to the Excel Sheet 2: \n The 'Program Data' " )
input_entry2 = tk.Entry(left_frame, text="", width=40)
browse2 = tk.Button(left_frame, text="Browse", command=input2)

input_path3 = tk.Label(left_frame, text="Link to Excel Sheet 3: \n The information about each district" )
input_entry3 = tk.Entry(left_frame, text="", width=40)
browse3 = tk.Button(left_frame, text="Browse", command=input3)

input_path_word_doc = tk.Label(left_frame, text="Link to the MailMerge Word Template" )
input_entry_word_doc = tk.Entry(left_frame, text="", width=40)
browse_word_doc = tk.Button(left_frame, text="Browse", command=input_word_doc)

output_path = tk.Label(right_frame, text="Where do you want your file saved? \n \n (I'd recommend a making a new folder in your desktop) \n")
output_entry = tk.Entry(right_frame, text="", width=40)
browse4 = tk.Button(right_frame, text="Browse", command=output)

begin_button = tk.Button(bottom_frame, text='Begin!', command = begin)

#%%
left_frame.pack(side=tk.LEFT)
#line.pack(pady=10)
right_frame.pack(side=tk.RIGHT)
bottom_frame.pack(side = tk.BOTTOM)

input_path1.pack(pady=5, padx = 45)
input_entry1.pack(pady=5, padx = 45)
browse1.pack(pady=5, padx = 45)

input_path2.pack(pady=5)
input_entry2.pack(pady=5)
browse2.pack(pady=5)

input_path3.pack(pady = 5)
input_entry3.pack(pady=5)
browse3.pack(pady=5)

input_path_word_doc.pack(pady = 5)
input_entry_word_doc.pack(pady=5)
browse_word_doc.pack(pady=5)

output_path.pack(pady=5, padx = 45)
output_entry.pack(pady=5, padx = 45)
browse4.pack(pady=5, padx = 45)

begin_button.pack(pady=20, padx = 25)

#%%
master.mainloop()

senators = pd.read_excel(input_entry1.get())
program_data = pd.read_excel(input_entry2.get())
mail_merge_template = pd.read_excel(input_entry2.get())
#source = os.path.abspath("C://Users//jwenger//Desktop//automated_reports//")

#%% methods:


#%%

