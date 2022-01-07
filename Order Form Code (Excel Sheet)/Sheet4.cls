VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub cmdDel_Click()

    Dim msgValue As VbMsgBoxResult

    msgValue = MsgBox("Do you want to delete the record?", vbYesNo + vbQuestion, "Delete")

    If msgValue = vbYes Then

        Call DeleteRecord

    End If

End Sub

Private Sub cmdModify_Click()

    Dim msgValue As VbMsgBoxResult

    msgValue = MsgBox("Do you want to modify the record?", vbYesNo + vbQuestion, "Modify")

    If msgValue = vbYes Then
        
        Call Reset

        Call Modify

    End If

End Sub

Private Sub cmdReset_Click()

    Dim msgValue As VbMsgBoxResult

    msgValue = MsgBox("Do you want to reset the Form?", vbYesNo + vbQuestion, "Reset")

    If msgValue = vbYes Then

        Call Reset

    End If

End Sub

Private Sub cmdSave_Click()

    If Validate = True Then

        Dim msgValue As VbMsgBoxResult

        msgValue = MsgBox("Do you want to save the data?", vbYesNo + vbQuestion, "Save")

          If msgValue = vbYes Then

                Call Save

                Call Reset
                

         End If

    End If
    

End Sub

