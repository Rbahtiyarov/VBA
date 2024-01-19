# Создание письма с темой и сообщением
Public Sub CreateMail()
Dim MyEmail As MailItem
' Create a new Outlook message item programatically
Set MyEmail = Application.CreateItem(olMailItem)
'Set your new message to, subject, body text and cc fields.
With MyEmail
'Komu
.To = "nnnn@mail.com"
'Subject
.Subject = "Ввести тему письма"
'Text
.Body = "Текст письма"

End With
MyEmail.Display
End Sub
