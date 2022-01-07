VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   7370
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   13280
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

