import win32com.client as win32
import time

outlook = win32.Dispatch("Outlook.Application")

addrs = [
    
]



anexo1 = r""
anexo2 = r""
anexo3 = r""

for addr in addrs:
    mail = outlook.CreateItem(0)
    mail.To = addr
    mail.Subject = ""

    # Obriga o Outlook a inserir a assinatura
    mail.Display()

    assinatura = mail.HTMLBody

    corpo = """
    <p></p>

    <p></p>


    <br>
    """

    mail.HTMLBody = corpo + assinatura
    mail.Attachments.Add(anexo1)
    mail.Attachments.Add(anexo2)
    mail.Attachments.Add(anexo2)

    mail.Send()
    print(f"Email enviado para {addr}")

    time.sleep(90)  # 1 e-mail a cada 1,5 minuto (IMPORTANTE)

print("Todos os e-mails foram enviados com sucesso")
