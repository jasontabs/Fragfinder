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
@sched.scheduled_job('interval', hours=1)
# defining timed job function
def timed_job():
    # subreddit to search
    subreddit = 'fragranceswap'

    # praw and reddit app stuff
    reddit = praw.Reddit(client_id='XXXX',
                         client_secret='XXXX',
                         user_agent='Frag Finder')

    # get 10 new posts from the fragranceswap subreddit
    new_posts = reddit.subreddit(subreddit).new(limit=10)
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
        print(fragrance)
        print(title)
        if fragrance:
            topics_dict["Title"].append(title)
            topics_dict["Flair"].append(flair)
            topics_dict["URL"].append(url)
            topics_dict["Body"].append(body)

            # data to pandas dataframe
            frag_data = pd.DataFrame(topics_dict)

            # sending dataframe to csv
            frag_data.to_csv('frags.csv', index=False)

            # Function to send the email
            def send_an_email():
                toaddr = 'XXXX@gmail.com'
                me = 'XXXXXXX@gmail.com'
                subject = "New Fragrance for Sale"

                msg = MIMEMultipart()
                msg['Subject'] = subject
                msg['From'] = me
                msg['To'] = toaddr
                msg.preamble = "New Fragrance Found"
                # msg.attach(MIMEText(text))

                part = MIMEBase('application', "octet-stream")
                part.set_payload(open("frags.csv", "rb").read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment; filename="frags.csv"')
                msg.attach(part)

                try:
                    s = smtplib.SMTP('smtp.gmail.com', 587)
                    s.ehlo()
                    s.starttls()
                    s.ehlo()
                    s.login(user='XXXXXX@gmail.com', password='XXXXXXXX')
                    # s.send_message(msg)
                    s.sendmail(me, toaddr, msg.as_string())
                    s.quit()
                # except:
                #   print ("Error: unable to send email")
                except smtplib.SMTPException as error:
                    print("Error")

            send_an_email()

sched.start()
