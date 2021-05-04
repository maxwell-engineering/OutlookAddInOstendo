' https://docs.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-your-first-vsto-add-in-for-outlook?redirectedfrom=MSDN&view=vs-2019
' https://docs.microsoft.com/en-us/visualstudio/vsto/programming-vsto-add-ins?view=vs-2019
' https://docs.microsoft.com/en-us/visualstudio/vsto/office-ui-customization?view=vs-2019
' https://docs.microsoft.com/en-us/visualstudio/vsto/how-to-add-commands-to-shortcut-menus?view=vs-2019

Imports Microsoft.Office.Core

Public Class ThisAddIn
    'Private WithEvents inspectors As Outlook.Inspectors

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon1()
    End Function

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        'Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup
        'inspectors = Application.Inspectors
    End Sub

    'Private Sub inspectors_NewInspector(ByVal Inspector As Microsoft.Office.Interop.Outlook.Inspector) Handles inspectors.NewInspector
    '    Dim mailItem As Outlook.MailItem = TryCast(Inspector.CurrentItem, Outlook.MailItem)
    '    If Not (mailItem Is Nothing) Then
    '        mailItem.Subject = "This text was added by using code"
    '        mailItem.Body = "This text was added by using code"
    '    End If
    'End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
