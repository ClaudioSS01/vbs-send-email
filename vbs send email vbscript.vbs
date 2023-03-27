Set objOutlook = CreateObject("Outlook.Application")
Set objMail = objOutlook.CreateItem(0)



objMail.To = "emiail@email"
objMail.Subject = "Assunto do Email"
objMail.Body = "Corpo do Email"
objMail.Attachments.Add "path\do\arquivo.txt"

objMail.Send

Set objMail = Nothing
Set objOutlook = Nothing
