# vbs-send-email
Ideal para usar com node.js no child_process para passar o comanendo

# Atenção
Esse codigo envia o email usando o outlook do windows que já deve estar instalado e usuario logado

```
Set objOutlook = CreateObject("Outlook.Application")
Set objMail = objOutlook.CreateItem(0)



objMail.To = "email@email"
objMail.Subject = "Assunto do Email"
objMail.Body = "Corpo do Email"
objMail.Attachments.Add "path\do\arquivo.txt"

objMail.Send

Set objMail = Nothing
Set objOutlook = Nothing

```
