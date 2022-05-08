import smtplib, ssl
from email.mime.text import MIMEText
import random
import time
import pandas as pd
import datetime
import json

excel_file = 'test-createReport (7) 220506.xls'
# Import settings in config.json
# Create a config.json file in the same folder with the following ...
#{   "smtp_server":"webmail.lnu.se",
#    "smtp_port":587,
#    "smtp_user": "xxxx@lnu.se",
#    "smtp_pass": "stefan",
#    "sender_email" "your@email.adress"
#}

with open('config.json') as config_file:
    config = json.load(config_file)

# There's a max limit on how many emails that can be sent every day, and you might want to limit the amount.
max_send_counter = 300
send_counter = 0

# Read data from xls file. Note, it is assumed that the first row is the header. And it should be a file that is exported from antagning.se, that is it should at least have the headers, "Efternamn"	"Förnamn", "Email". We are adding a new column "Sent" to the dataframe, if it doesn't exist yet all is added as a zero.

df = pd.read_excel(excel_file, index_col='Email')
# add sent column to dataframe
if not 'Sent' in df.columns:
    df.insert(2, 'Sent', 0)

# Create a secure SSL context

msg_body = """

Welcome to the course Introduction to Applied IoT at Linnaeus University.

The actual course starts at the 6th of June, but as you need to buy hardware in advance of the course we therefore contact you early on. That is why you are getting this email on your private mail, which will be an exception, and is needed because many have not yet gotten their LNU email in place. Further on you will get all information via the Discord server, or via Canvas (which will be automatically connected to your LNU email account).

What you need to do is:

(1): Check out the Welcome video: https://youtu.be/pEB_j0NFc74
(2): Join our Discord server: https://discord.gg/gfAKuC2Wfb
(3): Check the first HW Youtube video: https://youtu.be/5E3P1nlBHRs
(4): Buy appropiate hardware, you can find the links here: https://coursepress.lnu.se/courses/applied-iot/09-BOM

If there are any questions of any kind, please ask in the Discord server. Not via email or DM, keep all questions open. We will answer all questions as soon as possible.

I am so happy that you have joined this course, and I hope you will enjoy it.

Best,
Fredrik Ahlgren
Senior Lecturer, PhD

This message is sent from Python.
"""

# Try to log in to server and send email

def send_batch(df, config, msg_body, batch_size):
    batch_counter = 0
    global send_counter
    context = ssl.create_default_context()
    try:
        server = smtplib.SMTP(config["smtp_server"],config["smtp_port"])
        server.ehlo() # Can be omitted
        server.starttls(context=context) # Secure the connection
        server.ehlo() # Can be omitted
        server.login(config["smtp_user"], config["smtp_pass"])
        
        for index, row in df.iterrows():
            
            if row.Sent == 0:
                batch_counter += 1
                send_counter += 1
                if batch_counter >= batch_size:
                    print('... batch finished')
                    return
                if send_counter >= max_send_counter:
                    print('... max send counter reached')
                    return
                print('... sending mail '+ row.Förnamn.title() + ' ' + row.Efternamn.title())
                sleeptime = random.random()+0.1 # just to keep some randomness in here .. and to make it look more like real
                print('... sleeping for ' + str(sleeptime))
                time.sleep(sleeptime)
                msg = MIMEText('Hello ' + row.Förnamn.title() + ' ' + row.Efternamn.title() +','+ msg_body)
                msg['To'] = row.name
                msg['From'] = config['sender_email']
                msg['Subject'] = ('Welcome ' + row.Förnamn.title() + ' to Applied IoT 2022 at Linnaeus University')
                
                try:
                    server.sendmail(msg['From'], row.name , msg.as_string())
                    df.at[index, 'Sent'] = 1
                except Exception as e:
                    print(e)
                    df.at[index, 'Sent'] = 0
                    print('... failed to send mail to ' + row.name)
                    df.to_excel(excel_file)
                    #server.quit()
                    break
            else:
                print('... mail already sent to '+ row.Förnamn.title() + ' ' + row.Efternamn.title())

    except Exception as e:
        # Print any error messages to stdout
        print(e)
    finally:
        server.quit()
        df.to_excel(excel_file)

batch_size = 30

for i in range(0,max_send_counter//batch_size):
    send_batch(df, config, msg_body, batch_size)
    if send_counter >= max_send_counter:
        print('... max send counter reached')
        break