




Public Class Form1

    Dim app As Microsoft.Office.Interop.Outlook.Application

    Dim appNameSpace As Microsoft.Office.Interop.Outlook._NameSpace

    Dim memo As Microsoft.Office.Interop.Outlook.MailItem

    Dim outbox As Microsoft.Office.Interop.Outlook.MAPIFolder

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Try

            app = New Microsoft.Office.Interop.Outlook.Application
            appNameSpace = app.GetNamespace("MAPI")
            appNameSpace.Logon(Nothing, Nothing, False, False)

            memo = app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
            memo.To = "codyn@emeraldaire.com"
            memo.Subject = "subjectTest"
            memo.Body = "Hello there"
            memo.Send()

        Catch ex As Exception
            Console.WriteLine(ex.Message)

        End Try
    End Sub
End Class
