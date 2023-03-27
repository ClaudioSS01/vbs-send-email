# vbs-send-email
Ideal para usar com node.js no child_process para passar o comanendo

# Atenção
Esse codigo envia o email usando o outlook do windows que já deve estar instalado e usuario logado

```
Set objOutlook = CreateObject("Outlook.Application")
Set objMail = objOutlook.CreateItem(0)



objMail.To = "claudio.santos86@yahoo.com.br"
objMail.Subject = "Assunto do Email"
objMail.Body = "Corpo do Email"
objMail.Attachments.Add "C:\Users\Humanitar\Documents\CITRIX BOT\modelos relatorios\HSH.xls"

objMail.Send

Set objMail = Nothing
Set objOutlook = Nothing

```
