Attribute VB_Name = "efec"
' Every Form, Every Controls by Thomas Greenwood
' thomasgreenwood@2die4.com
'
' No credit Required to use.
'
' Known Limitations:
'   1.) On unload of a form all changes are reset
'       Solution: Call SetForms before loading the form.
'   2.) My knowledge of the On Error Statement and its non-obvious effects
'       are low so if you know, please let me know.
'
' To prevent controls you do not wish editing from been so, add an If statement to the
' For Each In Next Loops to check the .name of the control (no it won't remain as tempControl), or set the
' Propertys when the loop(s) is/are finished
'
' You can also use this statement to process the Type Of Control you wish to edit:
' If TypeOf Control = typeOfControl Then Control.Blah
' e.g. If TypeOf tempControl = TextBox Then tempControl.Text = "TextBox"
'
' I used this code to set the picture of every form to a dynamically loaded pic
'
' Thanks Very Much,
'
' Thomas Greenwood
Dim myForms As New Collection

Sub Main()
On Error GoTo EH
    myForms.Add frmMain ' Add every form you wish to be processed in the collection
    myForms.Add frmAbout
    SetForms ' The main routine, every form, every control
    SetTitles App.Title ' Set the title of the forms to app.title
    frmMain.Show 1 ' Show the main form to show what we have done
    End
Exit Sub
EH:
    Debug.Print Err.Description
    MsgBox Err.Description, vbOKOnly + vbExclamation + vbApplicationModal, "Added to immediate window."
End Sub

Sub SetForms()
On Error GoTo EH ' Show Errors
    Dim tempForm As Form
    Dim tempControl As Control
    
    ' Here I set every control in every form to a white forecolor
        For Each tempForm In myForms 'Every Form
            On Error Resume Next ' Ignore Errors
            tempForm.ForeColor = vbWhite
            tempForm.BackColor = &HE0E0E0
            For Each tempControl In tempForm.Controls 'Every Control On the form
                On Error Resume Next 'Ignore error because not every control will support the property
                tempControl.ForeColor = vbWhite ' Set some propertys
                tempControl.BackColor = &HE0E0E0
                tempControl.FontName = "Tahoma"
                tempControl.FontSize = 10
                tempControl.Style = 1
            Next tempControl ' The next control on the form
        Next tempForm ' The next Form
        
    On Error GoTo EH ' Show Errors - Dont know if putting this here does anything.
Exit Sub
EH:
    Debug.Print Err.Description
    MsgBox Err.Description, vbOKOnly + vbExclamation + vbApplicationModal, "Added to immediate window."
End Sub

Public Sub SetTitles(CaptionA As String) 'Called from frmMain and main()
' Here I set every form to the caption parameter
    On Error GoTo EH
    Dim tempForm As Form
    
        For Each tempForm In myForms 'Every Form
            On Error Resume Next
            tempForm.Caption = CaptionA
        Next tempForm ' The next Form
    On Error GoTo EH ' Show Errors - Dont know if putting this here does anything.
Exit Sub
EH:
    Debug.Print Err.Description
    MsgBox Err.Description, vbOKOnly + vbExclamation + vbApplicationModal, "Added to immediate window."
End Sub
