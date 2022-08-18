# %%
from os import name
import json
import smtplib
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from prompt_toolkit import prompt


### -----------------------------------------------------------------
### Variables
email = "hello@yourdomain.com"
pass_dir = open("./email_pass.json")
password = json.load(pass_dir)["pass"]

# message
subject = "Hello World!"
messageHTML = open("./email_template.html").read()  # your html email template
messagePlain = open("./email_template.txt").read()  # your txt email
# Google Sheet
SPREADSHEET_ID = "your Google Spreadsheet ID"
RANGE_NAME = "Name of your worksheet"
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

### -----------------------------------------------------------------
### Get GSheet to DF
# Authenticate service account
gc = gspread.service_account(filename="path to your service_account.json file")

# Open Sheet by key
sh = gc.open_by_key(SPREADSHEET_ID)

# Open Worksheet
worksheet = sh.worksheet(RANGE_NAME)

# create DF
df_raw = pd.DataFrame(worksheet.get_all_records())

# # filter by active stockists
# df_filter = df_raw.loc[df_raw['active'] == 'TRUE']
print(
    f"\n --- DETAILS--- \nEmail subject: {subject} \nrecipients: ",
    df_raw.loc[:, ("name", "email")],
)


# %%
### -----------------------------------------------------------------
# ### Loop over emails
a = input(
    "\n------- ATTENTION ------- \n1. Please review subject, names and recipients from above \n2. Enter 'yes' to confirm process: "
)
if a == "yes":
    for row in df_raw.itertuples():
        msg = MIMEMultipart("alternative")
        msg["From"] = email
        msg["To"] = row.email
        msg["Subject"] = subject
        # Attach message
        msg.attach(MIMEText(messagePlain.format(name=row.name), "plain"))
        msg.attach(MIMEText(messageHTML.format(name=row.name), "html"))
        ## Send Email
        server = smtplib.SMTP("smtp.dreamhost.com", 587)
        server.login(email, password)
        text = msg.as_string()
        server.sendmail(email, row.email, text)
        server.quit()
        print("email sent to {}".format(row.email))
else:
    print("Process cancelled")

print("\nProcess completed\n")

# %%
