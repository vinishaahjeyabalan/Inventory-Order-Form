Attribute VB_Name = "Module1"
Function Validate() As Boolean

    Dim frm As UserForm
    
    Set frm = UserForm1
    
    Validate = True
    
    With frm
    
'        reqBy.Interior.Color = xlNone
'        reqDate.Interior.Color = xlNone
'
'        prodGrp1.Interior.Color = xlNone
'        prodGrp2.Interior.Color = xlNone
'        prodGrp3.Interior.Color = xlNone
'        prodGrp4.Interior.Color = xlNone
'        prodGrp5.Interior.Color = xlNone
'
'        det1.Interior.Color = xlNone
'        det2.Interior.Color = xlNone
'        det3.Interior.Color = xlNone
'        det4.Interior.Color = xlNone
'        det5.Interior.Color = xlNone
'
'        qty1.Interior.Color = xlNone
'        qty2.Interior.Color = xlNone
'        qty3.Interior.Color = xlNone
'        qty4.Interior.Color = xlNone
'        qty5.Interior.Color = xlNone
'
'        prjName1.Interior.Color = xlNone
'        prjName2.Interior.Color = xlNone
'        prjName3.Interior.Color = xlNone
'        prjName4.Interior.Color = xlNone
'        prjName5.Interior.Color = xlNone
        
    End With
    
    'Validating Requested by
    
    If Trim(frm.reqBy.Text) = "" Then
        MsgBox "Requested by is blank.", vbOKOnly + vbInformation, "Requested by"
        'frm.reqBy.Select
        'reqBy.Interior.Color = vbRed
        Validate = False
        Exit Function
    End If
    
    
    'Validating Requested Date
    
    If Trim(frm.reqDate.Value) = "" Then
        MsgBox "Requested Date is blank.", vbOKOnly + vbInformation, "Requested Date"
        'frm.reqDate.Select
        'reqDate.Interior.Color = vbRed
        Validate = False
        Exit Function
    End If
    
    
    'Validating Product Group
    
    If Trim(frm.prodGrp1.Value) = "" Then
        MsgBox "Product Group is blank.", vbOKOnly + vbInformation, "Product Group"
        'frm.prodGrp1.Select
        'prodGrp1.Interior.Color = vbRed
        Validate = False
        Exit Function
    End If
    
    
    'Validating Details (functionality)
    
    If Trim(frm.det1.Value) = "" Then
        MsgBox "Details (functionality) is blank.", vbOKOnly + vbInformation, "Details (functionality)"
        'frm.det1.Select
        'det1.Interior.Color = vbRed
        Validate = False
        Exit Function
    End If
    
    
    'Validating Quantity
    
    If Trim(frm.qty1.Value) = "" Or Not IsNumeric(Trim(frm.qty1.Value)) Then
        MsgBox "Please enter valid Quantity.", vbOKOnly + vbInformation, "Quantity"
        'frm.qty1.Select
        'qty1.Interior.Color = vbRed
        Validate = False
        Exit Function
    End If
    
    
    'Validating Project Name
    
    If Trim(frm.prjName1.Value) = "" Then
        MsgBox "Project Name is blank.", vbOKOnly + vbInformation, "Project Name"
        'frm.prjName1.Select
        'prjName1.Interior.Color = vbRed
        Validate = False
        Exit Function
    End If
    
    
    
    
            
End Function



Sub Reset()

    Dim frm As UserForm
    
    Set frm = UserForm1

    With frm

        'reqBy.Interior.Color = xlNone
        frm.reqBy.Value = ""
        
        'reqDate.Interior.Color = xlNone
        frm.reqDate.Value = ""
        
        'prodGrp1.Interior.Color = xlNone
        frm.prodGrp1.Value = ""
        'prodGrp2.Interior.Color = xlNone
        frm.prodGrp2.Value = ""
        'prodGrp3.Interior.Color = xlNone
        frm.prodGrp3.Value = ""
        'prodGrp4.Interior.Color = xlNone
        frm.prodGrp4.Value = ""
        'prodGrp5.Interior.Color = xlNone
        frm.prodGrp5.Value = ""
        
        'det1.Interior.Color = xlNone
        frm.det1.Value = ""
        'det2.Interior.Color = xlNone
        frm.det2.Value = ""
        'det3.Interior.Color = xlNone
        frm.det3.Value = ""
        'det4.Interior.Color = xlNone
        frm.det4.Value = ""
        'det5.Interior.Color = xlNone
        frm.det5.Value = ""
        
        'qty1.Interior.Color = xlNone
        frm.qty1.Value = ""
        'qty2.Interior.Color = xlNone
        frm.qty2.Value = ""
        'qty3.Interior.Color = xlNone
        frm.qty3.Value = ""
        'qty4.Interior.Color = xlNone
        frm.qty4.Value = ""
        'qty5.Interior.Color = xlNone
        frm.qty5.Value = ""
        
        'prjName1.Interior.Color = xlNone
        frm.prjName1.Value = ""
        'prjName2.Interior.Color = xlNone
        frm.prjName2.Value = ""
        'prjName3.Interior.Color = xlNone
        frm.prjName3.Value = ""
        'prjName4.Interior.Color = xlNone
        frm.prjName4.Value = ""
        'prjName5.Interior.Color = xlNone
        frm.prjName5.Value = ""
        
    End With

End Sub



Sub Save()

    Dim frm As UserForm
    
    Dim database As Worksheet
    

    Dim iRow As Long

    Dim iSerial As Long
    
    Dim productLen As Integer
    
    Dim detailsLen As Integer
    
    Dim qtyLen As Integer
    
    Dim prjnameLen As Integer

   
    Set frm = UserForm1

    Set database = ThisWorkbook.Sheets("Database")
    
    
    Dim i As Integer
    
    Dim Products() As Variant
    If Trim(frm.prodGrp1.Value) <> "" Then
        ReDim Preserve Products(0)
        'ReDim Preserve Products(UBound(Products) + 1)
        Products(0) = frm.prodGrp1.Value
    End If
    
    If Trim(frm.prodGrp1.Value) <> "" And Trim(frm.prodGrp2.Value) <> "" Then
        ReDim Preserve Products(UBound(Products) + 1)
        Products(UBound(Products)) = frm.prodGrp2.Value
    End If
    
    If Trim(frm.prodGrp1.Value) <> "" And Trim(frm.prodGrp2.Value) <> "" And Trim(frm.prodGrp3.Value) <> "" Then
        ReDim Preserve Products(UBound(Products) + 1)
        Products(UBound(Products)) = frm.prodGrp3.Value
    End If
    
    If Trim(frm.prodGrp1.Value) <> "" And Trim(frm.prodGrp2.Value) <> "" And Trim(frm.prodGrp3.Value) <> "" And Trim(frm.prodGrp4.Value) <> "" Then
        ReDim Preserve Products(UBound(Products) + 1)
        Products(UBound(Products)) = frm.prodGrp4.Value
    End If
    
    If Trim(frm.prodGrp1.Value) <> "" And Trim(frm.prodGrp2.Value) <> "" And Trim(frm.prodGrp3.Value) <> "" And Trim(frm.prodGrp4.Value) <> "" And Trim(frm.prodGrp5.Value) <> "" Then
        ReDim Preserve Products(UBound(Products) + 1)
        Products(UBound(Products)) = frm.prodGrp5.Value
    End If
    
    
    Dim Details() As Variant
    If Trim(frm.det1.Value) <> "" Then
        ReDim Preserve Details(0)
        'ReDim Preserve Details(UBound(Details) + 1)
        Details(0) = frm.det1.Value
    End If
    
    If Trim(frm.det1.Value) <> "" And Trim(frm.det2.Value) <> "" Then
        ReDim Preserve Details(UBound(Details) + 1)
        Details(UBound(Details)) = frm.det2.Value
    End If
    
    If Trim(frm.det1.Value) <> "" And Trim(frm.det2.Value) <> "" And Trim(frm.det3.Value) <> "" Then
        ReDim Preserve Details(UBound(Details) + 1)
        Details(UBound(Details)) = frm.det3.Value
    End If
    
    If Trim(frm.det1.Value) <> "" And Trim(frm.det2.Value) <> "" And Trim(frm.det3.Value) <> "" And Trim(frm.det4.Value) <> "" Then
        ReDim Preserve Details(UBound(Details) + 1)
        Details(UBound(Details)) = det4.Value
    End If
    
    If Trim(frm.det1.Value) <> "" And Trim(frm.det2.Value) <> "" And Trim(frm.det3.Value) <> "" And Trim(frm.det4.Value) <> "" And Trim(frm.det5.Value) <> "" Then
        ReDim Preserve Details(UBound(Details) + 1)
        Details(UBound(Details)) = frm.det5.Value
    End If
    
    
    Dim Qty() As Variant
    If Trim(frm.qty1.Value) <> "" Then
        ReDim Preserve Qty(0)
        'ReDim Preserve Qty(UBound(Qty) + 1)
        Qty(0) = frm.qty1.Value
    End If
    
    If Trim(frm.qty1.Value) <> "" And Trim(frm.qty2.Value) <> "" Then
        ReDim Preserve Qty(UBound(Qty) + 1)
        Qty(UBound(Qty)) = frm.qty2.Value
    End If
    
    If Trim(frm.qty1.Value) <> "" And Trim(frm.qty2.Value) <> "" And Trim(frm.qty3.Value) <> "" Then
        ReDim Preserve Qty(UBound(Qty) + 1)
        Qty(UBound(Qty)) = frm.qty3.Value
    End If
    
    If Trim(frm.qty1.Value) <> "" And Trim(frm.qty2.Value) <> "" And Trim(frm.qty3.Value) <> "" And Trim(frm.qty4.Value) <> "" Then
        ReDim Preserve Qty(UBound(Qty) + 1)
        Qty(UBound(Qty)) = frm.qty4.Value
    End If
    
    If Trim(frm.qty1.Value) <> "" And Trim(frm.qty2.Value) <> "" And Trim(frm.qty3.Value) <> "" And Trim(frm.qty4.Value) <> "" And Trim(frm.qty5.Value) <> "" Then
        ReDim Preserve Qty(UBound(Qty) + 1)
        Qty(UBound(Qty)) = frm.qty5.Value
    End If
    
    
    Dim PrjName() As Variant
    If Trim(frm.prjName1.Value) <> "" Then
        ReDim Preserve PrjName(0)
        'ReDim Preserve PrjName(UBound(PrjName) + 1)
        PrjName(0) = frm.prjName1.Value
    End If
    
    If Trim(frm.prjName1.Value) <> "" And Trim(frm.prjName2.Value) <> "" Then
        ReDim Preserve PrjName(UBound(PrjName) + 1)
        PrjName(UBound(PrjName)) = frm.prjName2.Value
    End If
    
    If Trim(frm.prjName1.Value) <> "" And Trim(frm.prjName2.Value) <> "" And Trim(frm.prjName3.Value) <> "" Then
        ReDim Preserve PrjName(UBound(PrjName) + 1)
        PrjName(UBound(PrjName)) = frm.prjName3.Value
    End If
    
    If Trim(frm.prjName1.Value) <> "" And Trim(frm.prjName2.Value) <> "" And Trim(frm.prjName3.Value) <> "" And Trim(frm.prjName4.Value) <> "" Then
        ReDim Preserve PrjName(UBound(PrjName) + 1)
        PrjName(UBound(PrjName)) = frm.prjName4.Value
    End If
    
    If Trim(frm.prjName1.Value) <> "" And Trim(frm.prjName2.Value) <> "" And Trim(frm.prjName3.Value) <> "" And Trim(frm.prjName4.Value) <> "" And Trim(frm.prjName5.Value) <> "" Then
        ReDim Preserve PrjName(UBound(PrjName) + 1)
        PrjName(UBound(PrjName)) = frm.prjName5.Value
    End If

   
    If Trim(frm.TextBox2.Value) = "" Then
        iRow = database.Range("A" & Application.Rows.Count).End(xlUp).Row + 1
        
        If iRow = 2 Then
            iSerial = 1
        Else
            iSerial = database.Cells(iRow - 1, 1).Value + 1
            
        End If
        
    Else
        iRow = frm.TextBox1.Value
        iSerial = frm.TextBox2.Value
        
    End If
    
    
    Dim orderText As String
    Dim mailText As String
    
    productLen = (UBound(Products) - LBound(Products) + 1)
    detailsLen = (UBound(Details) - LBound(Details) + 1)
    qtyLen = (UBound(Qty) - LBound(Qty) + 1)
    prjnameLen = (UBound(PrjName) - LBound(PrjName) + 1)
    
    
    With database
    
        If (productLen = detailsLen And productLen = qtyLen And detailsLen = prjnameLen) Then
            For i = 0 To UBound(Products)
                .Cells(iRow, 1).Offset(i, 0).Value = iSerial
                iSerial = iSerial + 1
            
                .Cells(iRow, 2).Offset(i, 0).Value = frm.reqBy.Value
            
                .Cells(iRow, 3).Offset(i, 0).Value = frm.reqDate.Value
            
                .Cells(iRow, 4).Offset(i, 0).Value = Products(i)
            
                .Cells(iRow, 5).Offset(i, 0).Value = Details(i)
           
                .Cells(iRow, 6).Offset(i, 0).Value = Qty(i)
            
                .Cells(iRow, 7).Offset(i, 0).Value = PrjName(i)
            
                .Cells(iRow, 8).Offset(i, 0).Value = Application.UserName
                
                
                orderText = "Order ID: " & (iSerial - 1) & vbNewLine & _
                            "Requested by: " & frm.reqBy.Value & vbNewLine & _
                            "Requested Date: " & frm.reqDate.Value & vbNewLine & _
                            "Product Group: " & Products(i) & vbNewLine & _
                            "Details (functionality): " & Details(i) & vbNewLine & _
                            "Quantity: " & Qty(i) & vbNewLine & _
                            "Project Name: " & PrjName(i) & vbNewLine & vbNewLine


                mailText = mailText + orderText
            Next i
            
        Else
            MsgBox "The number of filled Product Name, Product Details, Quantity and Project Name are not tally", vbOKOnly + vbInformation, "Not Tally"
        
        End If
    
    End With
    
    
    If (productLen = detailsLen And productLen = qtyLen And detailsLen = prjnameLen) Then
    
        Dim xOutlookObj As Object
        Dim xOutApp As Object
        Dim xOutMail As Object
        Dim xMailBody As String
        On Error Resume Next
        Set xOutApp = CreateObject("Outlook.Application")
        Set xOutMail = xOutApp.CreateItem(0)
        xMailBody = Application.UserName & " responded to your form." & vbNewLine & vbNewLine & _
                    mailText & _
                    "Thank you"
                      On Error Resume Next
        With xOutMail
            .To = "vinishaah.jeyabalan@intel.com"
            .CC = ""
            .BCC = ""
            .Subject = "New Response on HSD Request Sample Order"
            .Body = xMailBody
            '.Attachments.Add ActiveWorkbook.FullName
            .Display   'or use .Send
        End With
        On Error GoTo 0
        Set xOutMail = Nothing
        Set xOutApp = Nothing
    
    End If
    
    
    frm.TextBox1.Value = ""
    frm.TextBox2.Value = ""
        
End Sub



Sub Modify()

    Dim frm As UserForm
    
    Set frm = UserForm1

    Dim iRow As Long
    Dim iSerial As Long
    
    iSerial = Application.InputBox("Please enter Serial Number to make modification.", "Modify", , , , , , 1)

    On Error Resume Next

    iRow = Application.WorksheetFunction.IfError _
    (Application.WorksheetFunction.Match(iSerial, Sheets("Database").Range("A:A"), 0), 0)
    
    On Error GoTo 0
    
    If iRow = 0 Then

         MsgBox "No record found.", vbOKOnly + vbCritical, "No Record"

        Exit Sub

    End If

    frm.TextBox1.Value = iRow
    frm.TextBox2.Value = iSerial
    
    frm.reqBy.Value = Sheets("Database").Cells(iRow, 2).Value
    frm.reqDate.Value = Sheets("Database").Cells(iRow, 3).Value
    frm.prodGrp1.Value = Sheets("Database").Cells(iRow, 4).Value
    frm.det1.Value = Sheets("Database").Cells(iRow, 5).Value
    frm.qty1.Value = Sheets("Database").Cells(iRow, 6).Value
    frm.prjName1.Value = Sheets("Database").Cells(iRow, 7).Value
    
End Sub



Sub DeleteRecord()

    Dim iRow As Long

    Dim iSerial As Long

    iSerial = Application.InputBox("Please enter S.No. to delete the recor.", "Delete", , , , , , 1)

    On Error Resume Next

    iRow = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Match(iSerial, Sheets("Database").Range("A:A"), 0), 0)

     On Error GoTo 0

    If iRow = 0 Then

        MsgBox "No record found.", vbOKOnly + vbCritical, "No Record"

        Exit Sub

    End If

   

    Sheets("Database").Cells(iRow, 1).EntireRow.Delete shift:=xlUp

End Sub





