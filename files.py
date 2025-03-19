import pandas as pd

data = {
    "Company": ["Google", "Microsoft", "Amazon"],
    "Email": ["hr@google.com", "hr@microsoft.com", "hiring@amazon.com"],
    "MailBody": [
        "Hi {company}, I am interested in a role at your company.",
        "Hello {company}, I would love to connect for an opportunity.",
        "Dear {company}, I’m reaching out regarding job openings."
    ]
}

df = pd.DataFrame(data)
df.to_excel("emails.xlsx", index=False)

print("Sample emails.xlsx file created!")
