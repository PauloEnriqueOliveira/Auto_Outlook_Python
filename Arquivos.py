import os
import win32com.client

SENDERS = ["emails para serem lidos"]
FOLDER = "caminho aonde serão salvos os arquivos"

if not os.path.isdir(FOLDER):
    os.makedirs(FOLDER)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)

messages = inbox.Items

unread_messages = messages.Restrict("[Unread] = True")

filtered_messages = []
for message in unread_messages:
    sender_email = message.SenderEmailAddress
    if sender_email in SENDERS:
        filtered_messages.append(message)

for message in filtered_messages:
    print(f"Processando email: {message.Subject}")
    attachments = message.Attachments
    attachment_count = attachments.Count
    print(f"Total de anexos encontrados: {attachment_count}")

    for i in range(1, attachment_count + 1):
        attachment = attachments.Item(i)
        print(f"Encontrado anexo: {attachment.FileName}")

        base_filename, file_extension = os.path.splitext(attachment.FileName)
        save_path = os.path.join(FOLDER, attachment.FileName)
        counter = 1

        while os.path.exists(save_path):
            save_path = os.path.join(FOLDER, f"{base_filename}_{counter}{file_extension}")
            counter += 1

        try:
            print(f"Salvando anexo em: {save_path}")
            attachment.SaveAsFile(save_path)
            print(f"Anexo salvo com sucesso: {save_path}")
            message.Unread = False 
            message.Save() 
        except Exception as e:
            print(f"Erro ao salvar o anexo: {e}")

print("Processamento concluído.")
