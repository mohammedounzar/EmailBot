import pandas as pd

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

sheet_cells.to_csv("C:/Users/mohamed/Downloads/Updated_cells_with_emails.csv", index=False)

print(sheet_cells['Email Address'])