import win32com.client as win32

def enviar_email_com_anexo(destinatario, copia, assunto, corpo, anexo):
    try:
        outlook = win32.Dispatch('Outlook.Application')
        email = outlook.CreateItem(0)
        email.To = destinatario
        email.CC = ";".join(copia)
        email.Subject = assunto
        email.Body = corpo
        email.Attachments.Add(anexo)
        email.Send()
        print(f"Email enviado para {destinatario} com cópia para {', '.join(copia)}.")
    except Exception as e:
        print(f"Erro ao enviar o email: {e}")
	
destinatario = "email do destinatario"
copia = ["usuarios em copia do email"]
assunto = "Assunto do email"
corpo = "Corpo do email"
anexo = 'Caminho do anexo'

enviar_email_com_anexo(destinatario, copia, assunto, corpo, anexo)
