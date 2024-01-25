#Python code to handle mentor matching emails.
import win32com.client as win32
import pandas as pd
import time

#Set up outlook to run a test
outlook = win32.Dispatch('outlook.application')

#Add preadsheet to document
file_path = r'Matching Sheet.xlsx'
#read the correct spreadsheet
df1 = pd.read_excel(file_path, sheet_name = 'Sheet1')

print("Please input the cohort name in the format period - season i.e. 23-24 – Winter")
cohort_name = input("Enter cohort name:")

#get a list on the mentees names
listmentees = df1['Index'].unique()
#calculate how many mentees are on the list
list_carrier = list(listmentees)

##List the mentess so i know that part worked.
print(list_carrier)

index = 0
for i in list_carrier:
    
    index = index + 1
    #Limited to 250 because outlook only allows someone to send 300 emails per day.
    #I wanted to leave space to conduct normal email business. 
    if(index>250):
        print("Limit exceeded")
        break

    #Break each database line down into seperate variables.
    mentee_email = df1.loc[df1['Index'] == i, 'Mentee Email'].item()
    mentee_name = df1.loc[df1['Index'] == i, 'Mentee Name'].item()
    mentor_email = df1.loc[df1['Index'] == i, 'Mentor Email'].item()
    mentor_name = df1.loc[df1['Index'] == i, 'Mentor Name'].item()

    #Print the names out for testing

    #Prepare the mentee email
    mail = outlook.CreateItem(0)
    email = mentee_email
    mail.cc = '###'
    mail.To = email
    mail.Subject = "HDN Staff Mentoring Programme - Matching you with your Mentor"
    outputstring = """\
    Dear {}, 

    We are delighted to let you know that you have now been matched with a mentor on the Staff Mentoring Programme {} Cohort.
    Your mentor’s name is {}
    You can contact them on {}


    We recommend that the mentee makes first contact, so we recommend you get in touch as soon as possible.
    Please contact us at #### if you have any questions.
    We look forward to working with you during the programme.
    From the HDN Mentoring Team
    
    """
    mail.Body = outputstring.format(mentee_name,cohort_name,mentor_name,mentor_email);
    #path_1=path+i+".xlsx"
    #mail.Attachments.Add(path_1)
    #print(mail.Body)
    
    printtxt = "Sending email {} out of {}"
    print(printtxt.format(index,len(list_carrier)*2))
    mail.Send()
    print("Mail sent")

    time.sleep(10)

    index = index + 1
    #preparing the mentor email
    mail = outlook.CreateItem(0)
    email = mentor_email
    mail.cc = '###'
    mail.To = email
    mail.Subject = "HDN Staff Mentoring Programme - Matching you with your Mentee"
    outputstring = """\
    Dear {}, 

    We are delighted to let you know that you have now been matched with a mentee on the Staff Mentoring Programme {} Cohort.
    Your mentee’s name is {}
    You can contact them on {}


    We recommend that the mentee makes first contact, but if your mentee does not get in touch soon please get in touch with them.
    Please contact us at ### if you have any questions.
    We look forward to working with you during the programme.
    From the HDN Mentoring Team
    
    """
    mail.Body = outputstring.format(mentor_name,cohort_name,mentee_name,mentee_email);
    #path_1=path+i+".xlsx"
    #mail.Attachments.Add(path_1)
    #print(mail.Body)
    
    printtxt = "Sending email {} out of {}"
    print(printtxt.format(index,len(list_carrier)*2))
    mail.Send()
    print("Mail sent")
    
    time.sleep(10)
