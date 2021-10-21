from __future__ import print_function
import os
import random as r
import xml.etree.ElementTree as ET
from tkinter import*
from tkinter import filedialog, messagebox
from PIL import ImageTk, Image
from mailmerge import MailMerge
from datetime import datetime


def format_xml(report):
    report = open(report, 'r')
    tree = ET.fromstring(report.read().strip())
    combined = []
    oth = False
    lead = False
    train = False
    for i in tree:  # This is objectively a terrible way of doing this, but it works and looking up documentation for
        # how to properly do it would take up too much time. This is not a super complex project so it is okay.
        for j in i:
            for k in j:
                for l in k:
                    for cell in l:
                        if cell.text is not None:
                            # print(cell.text)
                            if "Back of House" in cell.text:
                                combined.append("Back of House")
                            elif "Front of House" in cell.text:
                                combined.append("Front of House")
                            elif "Leadership" in cell.text:
                                combined.append("Leadership")
                                lead = True
                            elif "Other" in cell.text:
                                combined.append("Other")
                                oth = True
                            elif "Training" in cell.text:
                                combined.append("Training")
                                train = True
                            elif "Total Day" in cell.text or "Logbook Notes" in cell.text:
                                pass
                            else:
                                combined.append(cell.text)

    BOH = combined[combined.index("Back of House"):combined.index("Front of House")]
    leadership = []
    other = []
    training = []
    if lead:
        FOH = combined[combined.index("Front of House"):combined.index("Leadership")]
        if oth:
            leadership = combined[combined.index("Leadership"):combined.index("Other")]
            if train:
                other = combined[combined.index("Other"):combined.index("Training")]
                training = combined[combined.index("Training"):]
            else:
                other = combined[combined.index("Other"):]
        elif train:
            leadership = combined[combined.index("Leadership"):combined.index("Training")]
            training = combined[combined.index("Training"):]
        else:
            leadership = combined[combined.index("Leadership"):]
    elif oth:
        FOH = combined[combined.index("Front of House"):combined.index("Other")]
        if train:
            other = combined[combined.index("Other"):combined.index("Training")]
        else:
            other = combined[combined.index("Other"):]
    elif train:
        FOH = combined[combined.index("Front of House"):combined.index("Training")]
        training = combined[combined.index("Training"):]
    else:
        FOH = combined[combined.index("Front of House"):]
    # print(FOH)
    # print(BOH)
    # print(leadership)
    # print(training)
    # print(other)
    return [FOH, BOH, leadership, training, other]


def chunks(big_list):
    index = 1
    small_list = []
    while index < len(big_list)-1:
        counter = 0
        times = 0
        while times < 1:
            if ":" in big_list[index+counter]:  # This tests if it is a time or not, if there is a colon that implies
                # it is a time
                times += 1
                counter += 1
            counter += 1
        if counter == 4:
            small_list.append(big_list[index+1:index+counter])
        else:
            small_list.append(big_list[index:index+counter])
        index += counter
    if len(big_list) > 0:
        small_list.insert(0, big_list[0])
    else:
        small_list.append('null')
    return small_list

    # final_list = []
    # small_list = [big_list[x:x+4] for x in range(1, len(big_list), 4)]
    # for person in small_list:
    #     if person[1] == 'null':
    #         person.remove('null')
    #         final_list.append(person)
    #     else:
    #         final_list.append(person[1:])
    # return final_list


def sort_everything(schedule):
    split_times = []
    breaks = []
    no_breaks = []
    combined = []
    hours_worked = []
    for person in schedule[1:]:  # Skips the team "BOH" or "FOH
        if type(person) is list:
            # print(person)
            start_hour = float(person[1].split(':')[0])
            start_minute = float(person[1].split(':')[1][:2]) / 60  # The [:2] is to remove the am or pm
            # print(person[1].split(':')[1][:2])
            # print(person[1].split(':')[1][2:])
            end_hour = float(person[2].split(':')[0])
            end_minute = float(person[2].split(':')[1][:2]) / 60
            # print(person[2].split(':')[1][:2])
            # print(person[2].split(':')[1][2:])
            start = start_hour + start_minute
            end = end_hour + end_minute
            if (person[1].split(':')[1][2:] == 'am' or start >= 12) and person[2].split(':')[1][2:] == 'pm' and end < 12:
                #  The last two conditions catch any shifts that end before 1pm, those we want to calculate using the
                #  other method (the simple end - start) It is only for shifts that start in am and end in pm that we
                #  must use this method of 12 - start + end
                hours_worked.append(12 - start + end)
                if 12 - start + end >= 5:
                    no_breaks.append('')
                else:
                    no_breaks.append('----------------')
            else: # This should trigger if the team member is working at night or if
                hours_worked.append(end - start)
                if end - start >= 5:
                    no_breaks.append('')
                else:
                    no_breaks.append('----------------')
        else:
            hours_worked.append('')
            if person[0] == 'null':
                no_breaks.append('')
            else:
                no_breaks.append('----------------')
    for person, brk, hour in zip(schedule[1:], no_breaks, hours_worked):
        if type(person) is list:
            combined.append({'Team': person[0], 'Shift': person[1] + ' ' + person[2][0:-2] + ' (' + str(hour) + ')', 'nobrek': brk})
            # The [0:-2] removes the am or pm from the second number. person[1] refers to time 1 and person[2] to time 2
        else:
            if person == 'null':
                combined.append({'Team': '', 'Shift': '', 'nobrek': ''})
            else:
                combined.append({'Team': '', 'Shift': '', 'nobrek': ''})  # Provides a blank space before each
                combined.append({'Team': person, 'Shift': '----------------------', 'nobrek': brk})
    rows_needed = 49 - len(combined)
    if rows_needed > 0:
        for i in range(0, rows_needed):
            combined.append({'Team': '', 'Shift': '', 'nobrek': ''})
    return combined


# os.rename(r'report.xls', r'file2.xml')
# source = open("file2.xml")
# combined = format_xml(source)
# print(combined[0])
# print(combined[1])
# print(combined[2])
# print(combined[3])
# print(combined[4])
# FOH_final = chunks(combined[0])
# BOH_final = chunks(combined[1])
# leadership_final = chunks(combined[2])
# training_final = chunks(combined[3])
# other_final = chunks(combined[4])
#
# print(FOH_final)
# print(BOH_final)
# print(leadership_final)
# print(training_final)
# print(other_final)
# loc = "C:/Users/georg/Downloads"
# wb = xlrd.open_workbook("test.xls")
# sheet = wb.sheet_by_index(0)
# print(sheet.cell_value(0, 0)) this is code for excel


def open_file():
    root.name = filedialog.askopenfilename(initialdir="Desktop", title="Select A File",
                                           filetypes=[('excel files', '.xls')])
    try:
        os.rename(root.name, r'file.xml')
        root.file = r'file.xml'
    except: # this will try to rename the file to something else with a random number
        try:
            random = str(r.randint(1, 100000))
            rename = r'file' + random +'.xml'
            os.rename(root.name, rename)
            root.file = rename
        except:
            messagebox.showerror("Error", "Please remove any xml files in the folder and try again. If you just pressed cancel then do not worry about this message")


def write_date(entry):
    root.date = e.get()
    print(e.get())


def get_combined():
    it_worked = False
    combined = []
    try:
        combined = format_xml(root.file)
        it_worked = True
    except AttributeError:
        messagebox.showerror("Error", "Please select a valid excel file")
    if it_worked:
        try:
            FOH = chunks(combined[0])
            BOH = chunks(combined[1])
            leadership_final = chunks(combined[2])
            training_final = chunks(combined[3])
            other_final = chunks(combined[4])
            template = "files/shift.docx"
            FOH_final = FOH + leadership_final + training_final + other_final
            BOH_final = BOH + leadership_final + training_final + other_final
            actual_foh_final = sort_everything(FOH_final)
            actual_boh_final = sort_everything(BOH_final)
            print(actual_boh_final)
            print(actual_foh_final)
            FOH_mail = MailMerge(template)
            FOH_mail.merge_rows('Shift', actual_foh_final)
            FOH_mail.merge_rows('date', {'Team': '', 'Shift': '', 'nobrek': '', 'test123': datetime.today().strftime('%Y-%m-%d')})
            FOH_mail.write('FOH ' + root.date + '.docx')
            BOH_mail = MailMerge(template)
            BOH_mail.merge_rows('Shift', actual_boh_final)
            BOH_mail.merge_rows('date', {'Team': '', 'Shift': '', 'nobrek': '', 'test123': datetime.today().strftime('%Y-%m-%d')})
            BOH_mail.write('BOH ' + root.date + '.docx')
            messagebox.showinfo("Success!", "You may now exit the program")
        except:
            messagebox.showerror("Error", "List error, make sure you have BOH, FOH, Leadership, Training, and Other selected as well as in times, out times, first names, and prefered names")
    os.remove(root.file)

root = Tk()
root.title("Chick-fil-A Roster Report")
root.iconbitmap('files/chick.ico')
app_width = 400
app_height = 400
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = screen_width/2 - app_width/2  # This is how we center the app
y = screen_height/2 - app_height/2
root.geometry(f'{app_width}x{app_height}+{int(x)}+{int(y)}')


background_image = ImageTk.PhotoImage(Image.open("files/chick3.png"))
background_label = Label(root, image=background_image)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

e = Entry(root, width=20)
e.insert(0, "Enter Date for Roster")
e.grid(row=0, column=0, pady=60)
e.bind("<Return>", write_date)

names_btn = Button(root, text="Select Excel File", command=open_file, fg='blue').grid(row=1, column=0, pady=0, padx=60)
combine_btn = Button(root, text="Generate Roster", command=get_combined).grid(row=2, column=0, pady=60, padx=60)
button_exit = Button(root, text="Exit Program", command=root.quit, fg='red').grid(row=2, column=1, padx=30)
root.mainloop()

#  pyinstaller --onefile chick-fil-a.py --windowed --icon=C:\Users\Wesley\PycharmProjects\pythonProject\files\chick.ico
#  That is the the command line in terminal to convert this to an exe
