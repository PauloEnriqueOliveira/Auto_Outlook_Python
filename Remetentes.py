import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)

print(f"Nome da pasta: {inbox.Name}")

messages = inbox.Items

unread_messages = messages.Restrict("[Unread] = True")

for message in unread_messages:
    sender_email = message.SenderEmailAddress
    print(f"Remetente do email: {sender_email}") 
