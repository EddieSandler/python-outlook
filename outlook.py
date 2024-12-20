import win32com.client as client
outlook=client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# message=outlook.CreateItem(0)
# message.Display()
# message.To = "ed.sandler@gmail.com"
# message.Subject="Testing"
# message.Body="This works.... I think"
# message.Save()
# message.Send()

inbox=outlook.GetDefaultFolder(6)
messages =inbox.Items

for i, message in enumerate(messages):
    try:
        # Get the sender's email address or name
        sender = message.SenderEmailAddress  # or message.SenderName
        name=message.SenderName
        if name =="Braxton Bell":
            # print(f"Sender {i+1}: {sender}")
            print(message.Subject)
    except AttributeError:
        # Some items might not have the expected attributes
        print(f"Message {i+1} does not have a sender.")
