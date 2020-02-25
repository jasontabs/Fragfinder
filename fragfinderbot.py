#!/usr/bin/env python3
# scraping/parsing stuff
import praw
import re
import pandas as pd
# email stuffs
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
# job scheduler
from apscheduler.schedulers.blocking import BlockingScheduler

# initiate job scheduler
sched = BlockingScheduler()

# job runs every 1 hour
@sched.scheduled_job('interval', minutes=1)
# defining timed job function
def timed_job():
    # subreddit to search
    subreddit = 'fragranceswap'

    # praw and reddit app stuff
    reddit = praw.Reddit(client_id='XXXXXX',
                         client_secret='XXXXXX',
                         user_agent='XXXXXX')

    # get 10 new posts from the fragranceswap subreddit
    new_posts = reddit.subreddit(subreddit).new(limit=5)
    topics_dict = {
        "Title": [],
        "Flair": [],
        "URL": [],
        "Body": []}

    for post in new_posts:
        title = post.title
        flair = post.link_flair_text
        url = post.url
        body = post.selftext
        # This returns a Match Object. So, it doesn't return a string.
        # It returns a bool (T/F) so essentially my IF is asking if true
        # then add to dict.
        fragrance = re.search(r'le labo', title and body, re.I) and flair == 'SELLING'
        if fragrance:
            topics_dict["Title"].append(title)
            topics_dict["Flair"].append(flair)
            topics_dict["URL"].append(url)
            topics_dict["Body"].append(body)

    # data to pandas dataframe
    frag_data = pd.DataFrame(topics_dict)

    # check if the dataframe has new entries
    if not frag_data.empty:
        # sending dataframe to csv
        frag_data.to_excel('frags.xlsx', index=False)

        #formatting xlsx
        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter("frags.xlsx", engine='xlsxwriter')

        # Convert the dataframe to an XlsxWriter Excel object.
        frag_data.to_excel(writer, sheet_name='Sheet1', index=False, startrow=0)

        # Get the xlsxwriter workbook and worksheet objects.
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Add some cell formats.
        wrap = workbook.add_format({'text_wrap': True})
        #header format
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'align': 'center',
            'valign': 'top',
            'fg_color': '#969696',
            'border': 1})

        # Write the column headers with the defined format.
        for col_num, value in enumerate(frag_data.columns.values):
            worksheet.write(0, col_num + 0, value, header_format)


        # Set the column width and format.
        worksheet.set_column('A:D', None, wrap)
        worksheet.set_column('A:A', 50, None)
        worksheet.set_column('B:B', 10, None)
        worksheet.set_column('C:C', 50, wrap)
        worksheet.set_column('D:D', 100, None)




        # Close the Pandas Excel writer and output the Excel file.
        writer.save()




        # Function to send the email
        def send_an_email():
            toaddr = 'XXXXXXgmail.com'
            me = 'XXXXXX@gmail.com'
            subject = "New Fragrance for Sale"

            msg = MIMEMultipart()
            msg['Subject'] = subject
            msg['From'] = me
            msg['To'] = toaddr
            msg.preamble = "New Fragrance Found"
            # msg.attach(MIMEText(text))

            part = MIMEBase('application', "octet-stream")
            part.set_payload(open("frags.xlsx", "rb").read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="frags.xlsx"')
            msg.attach(part)

            try:
                s = smtplib.SMTP('smtp.gmail.com', 587)
                s.ehlo()
                s.starttls()
                s.ehlo()
                s.login(user='XXXXXX', password='XXXXX')
                # s.send_message(msg)
                s.sendmail(me, toaddr, msg.as_string())
                s.quit()
            # except:
            #   print ("Error: unable to send email")
            except smtplib.SMTPException as error:
                print("Error")

        send_an_email()

sched.start()
