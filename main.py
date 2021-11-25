import pandas as p
import smtplib as sm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

data = p.read_excel("mail1.xlsx")

print(type(data))

email_col = data.get("email")

list_of_emails = list(email_col)

print(list_of_emails)

try:
   server = sm.SMTP("smtp.gmail.com", 587)
   server.starttls()
   server.login("yourmsil@gmail.com", "your-password")
   from_ = "your-mail@gmail.com"
   to_ = list_of_emails
   message = MIMEMultipart("alternative")
   message['subject'] = "Subject"
   message["from"] = "your-mail@gmail.com"

   html='''
   <html>
   <head>
   
   </head>
   <body>

   <h1>Hello there,</h1>
   <h2>

 It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.

Why do we use it?
It is a long established fact that a reader will be distracted by the readable content of a page when looking at its layout. The point of using Lorem Ipsum is that it has a more-or-less normal distribution of letters, as opposed to using 'Content here, content here', making it look like readable English. Many desktop publishing packages and web page editors now use Lorem Ipsum as their default model text, and a search for 'lorem ipsum' will uncover many web sites still in their infancy. Various versions have evolved over the years, sometimes by accident, sometimes on purpose (injected humour and the like).




</h2>
   </body>
   </html>
   '''
   text = MIMEText(html, "html")

   message.attach(text)

   server.sendmail(from_, to_, message.as_string())
   print(("msg has successfully sent"))

except Exception as e:
    print(e)
