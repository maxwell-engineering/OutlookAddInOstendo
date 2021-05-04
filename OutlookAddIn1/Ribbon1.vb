'TODO:  Follow these steps to enable the Ribbon (XML) item:

'1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Ribbon1()
'End Function

'2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
'   actions, such as clicking a button. Note: if you have exported this Ribbon from the
'   Ribbon designer, move your code from the event handlers to the callback methods and
'   modify the code to work with the Ribbon extensibility (RibbonX) programming model.

'3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

'For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.

' https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ee692172(v=office.14)?redirectedfrom=MSDN#attachment-context-menu

<Runtime.InteropServices.ComVisible(True)> _
Public Class Ribbon1
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("OutlookAddIn1.Ribbon1.xml")
    End Function

#Region "Ribbon Callbacks"
    'Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
    Dim attSelection As Outlook.AttachmentSelection
    Dim selection As Outlook.Selection
    Dim mItem As Outlook.MailItem

    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub
    Public Sub GetButtonID(ByVal control As Office.IRibbonControl)
        Select Case TypeName(control.Context)
            Case "Selection"
                selection = control.Context
                mItem = selection.Item(1)
                'MsgBox("Menu selected " & control.Id & ":" & TypeName(selection.Item(1)), vbOK, "Right Click")
                MsgBox(mItem.SenderEmailAddress & ":" & mItem.Subject, vbOK, "Right Click")
            Case "AttachmentSelection"
                attSelection = control.Context
                mItem = attSelection.Parent
                MsgBox("Menu selected " & control.Id & ":" & attSelection.Item(1).DisplayName, vbOK, "Right Click")
                MsgBox(mItem.SenderEmailAddress & ":" & mItem.Subject, vbOK, "Right Click")
        End Select
    End Sub


#End Region

#Region "Helpers"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
