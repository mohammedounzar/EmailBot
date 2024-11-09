import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from getpass import getpass
import pandas as pd
'''
sheet_form = pd.read_csv("C:/Users/mohamed/Downloads/Copy of IEEE ENSIAS SB recruitment form (Responses) - Form Responses 1.csv")
sheet_cells = pd.read_csv("C:/Users/mohamed/Downloads/Copy of IEEE ENSIAS SB recruitment form (Responses) - cells.csv")

names_form = sheet_form['Full name'].tolist()
emails = sheet_form['Email Address'].tolist()  
names_cells = sheet_cells['Full Name'].tolist()

names_form = [str(name).strip() if isinstance(name, str) else "" for name in names_form]
names_cells = [str(name).strip() if isinstance(name, str) else "" for name in names_cells]
emails = [str(email).strip() if isinstance(email, str) else "" for email in emails]

name_email_dict = dict(zip(names_form, emails))

receiver_email = [name_email_dict.get(name, "") for name in names_cells]

sheet_cells['Email Address'] = receiver_email

sheet_cells.to_csv("C:/Users/mohamed/Downloads/Updated_cells.csv", index=False)

'''
sender_email = "doukhouhajar@ieee.org"
password = getpass("Enter your email password: ")

sheet = pd.read_csv("C:/Users/mohamed/Downloads/Updated_cells.csv")
receiver_email = sheet["Email Address"].tolist()
receiver_names = sheet["Full Name"].tolist()

print(receiver_email, receiver_names)
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(sender_email, password) 

for i in range(len(receiver_names)):
    subject = "IEEE ENSIAS Student Branch"
    html = f"""
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    background-color: #f4f4f4;
                    padding: 20px;
                    margin: 0;
                }}
                .container {{
                    background-color: #ffffff;
                    padding: 20px;
                    margin: auto;
                    width: 80%;
                    max-width: 600px;
                    border-radius: 8px;
                    box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
                }}
                .header {{
                    text-align: center;
                    margin-bottom: 20px;
                }}
                .header img {{
                    width: 150px;
                    margin-bottom: 10px;
                }}
                .content {{
                    text-align: center;
                }}
                .greeting {{
                    font-size: 24px;
                    font-weight: bold;
                    margin-bottom: 20px;
                    color: #036eb6;
                }}
                .message {{
                    font-size: 16px;
                    line-height: 1.5;
                    margin-bottom: 20px;
                    color: #036eb6;
                    white-space: normal;
                }}
                .footer {{
                    font-size: 14px;
                    color: #888;
                    text-align: center;
                    margin-top: 30px;
                }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <img src="https://raw.githubusercontent.com/khaouitiabdelhakim/ArabicOCR-Python-Tutorial/main/ieee.png" alt="IEEE ENSIAS Logo">
                </div>
                <div class="content">
                    <div class="greeting">Welcome, {receiver_names[i]}!</div>
                    <div class="message">
                        <p>Congratulations! You have been elected to be a member of the <strong>IEEE ENSIAS Family</strong>!</p>
                        <p>We are delighted to welcome you and look forward to seeing your active participation in our community.</p>
                    </div>
                </div>
                <div class="footer">
                    <p>Best Regards,<br>IEEE ENSIAS Student Branch</p>
                </div>
            </div>
        </body>
        </html>
    """

    message = MIMEMultipart("alternative")
    message["Subject"] = "Welcome to IEEE ENSIAS Student Branch"
    message["From"] = "IEEE ENSIAS Student Branch"
    message["To"] = receiver_email[i]
    message["Importance"] = "high"

    # Attach the body with the msg instance as HTML
    message.attach(MIMEText(html, 'html'))

    try:
        # Convert the message to a string and send the email
        text = message.as_string()
        server.sendmail(sender_email, receiver_email[i], text)
        print(f"Email sent successfully to {receiver_email[i]}!")

    except Exception as e:
        print(f"Failed to send email to {receiver_email[i]}. Error: {str(e)}")

# Close the SMTP session
server.quit()