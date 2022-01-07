Attribute VB_Name = "Module1"
Function Validate() As Boolean

    Dim frm As Worksheet
    
    Set frm = ThisWorkbook.Sheets("Form")
    
    Validate = True
    
    With frm
    
        .Range("G6").Interior.Color = xlNone
        .Range("G8").Interior.Color = xlNone
        
        .Range("G10").Interior.Color = xlNone
        .Range("G12").Interior.Color = xlNone
        .Range("G14").Interior.Color = xlNone
        .Range("G16").Interior.Color = xlNone
        .Range("G18").Interior.Color = xlNone
        
        .Range("G20").Interior.Color = xlNone
        .Range("G22").Interior.Color = xlNone
        .Range("G24").Interior.Color = xlNone
        .Range("G26").Interior.Color = xlNone
        .Range("G28").Interior.Color = xlNone
        
        .Range("G30").Interior.Color = xlNone
        .Range("G32").Interior.Color = xlNone
        .Range("G34").Interior.Color = xlNone
        .Range("G36").Interior.Color = xlNone
        .Range("G38").Interior.Color = xlNone
        
        .Range("G40").Interior.Color = xlNone
        .Range("G42").Interior.Color = xlNone
        .Range("G44").Interior.Color = xlNone
        .Range("G46").Interior.Color = xlNone
        .Range("G48").Interior.Color = xlNone
        
    End With
    
    'Validating Requested by
    
    If Trim(frm.Range("G6").Value) = "" Then
        MsgBox "Requested by is blank.", vbOKOnly + vbInformation, "Requested by"
        frm.Range("G6").Select
        frm.Range("G6").Interior.Color = vbRed
        Validate = False
        Exit Function
    End If
    
    
    'Validating Requested Date
    
    If Trim(frm.Range("G8").Value) = "" Then
        MsgBox "Requested Date is blank.", vbOKOnly + vbInformation, "Requested Date"
        frm.Range("G8").Select
        frm.Range("G8").Interior.Color = vbRed
        Validate = False
        Exit Function
    End If
    
    
    'Validating Product Group
    
    If Trim(frm.Range("G10").Value) = "" Then
        MsgBox "Product Group is blank.", vbOKOnly + vbInformation, "Product Group"
        frm.Range("G10").Select
        frm.Range("G10").Interior.Color = vbRed
        Validate = False
        Exit Function
    End If
    
    
    'Validating Details (functionality)
    
    If Trim(frm.Range("G20").Value) = "" Then
        MsgBox "Details (functionality) is blank.", vbOKOnly + vbInformation, "Details (functionality)"
        frm.Range("G20").Select
        frm.Range("G20").Interior.Color = vbRed
        Validate = False
        Exit Function
    End If
    
    
    'Validating Quantity
    
    If Trim(frm.Range("G30").Value) = "" Or Not IsNumeric(Trim(frm.Range("G30").Value)) Then
        MsgBox "Please enter valid Quantity.", vbOKOnly + vbInformation, "Quantity"
        frm.Range("G30").Select
        frm.Range("G30").Interior.Color = vbRed
        Validate = False
        Exit Function
    End If
    
    
    'Validating Project Name
    
    If Trim(frm.Range("G40").Value) = "" Then
        MsgBox "Project Name is blank.", vbOKOnly + vbInformation, "Project Name"
        frm.Range("G40").Select
        frm.Range("G40").Interior.Color = vbRed
        Validate = False
        Exit Function
    End If
    
    
    
    
            
End Function



Sub Reset()

    With Sheets("Form")

        .Range("G6").Interior.Color = xlNone
        .Range("G6").Value = ""
        
        .Range("G8").Interior.Color = xlNone
        .Range("G8").Value = ""
        
        .Range("G10").Interior.Color = xlNone
        .Range("G10").Value = ""
        .Range("G12").Interior.Color = xlNone
        .Range("G12").Value = ""
        .Range("G14").Interior.Color = xlNone
        .Range("G14").Value = ""
        .Range("G16").Interior.Color = xlNone
        .Range("G16").Value = ""
        .Range("G18").Interior.Color = xlNone
        .Range("G18").Value = ""
        
        .Range("G20").Interior.Color = xlNone
        .Range("G20").Value = ""
        .Range("G22").Interior.Color = xlNone
        .Range("G22").Value = ""
        .Range("G24").Interior.Color = xlNone
        .Range("G24").Value = ""
        .Range("G26").Interior.Color = xlNone
        .Range("G26").Value = ""
        .Range("G28").Interior.Color = xlNone
        .Range("G28").Value = ""
        
        .Range("G30").Interior.Color = xlNone
        .Range("G30").Value = ""
        .Range("G32").Interior.Color = xlNone
        .Range("G32").Value = ""
        .Range("G34").Interior.Color = xlNone
        .Range("G34").Value = ""
        .Range("G36").Interior.Color = xlNone
        .Range("G36").Value = ""
        .Range("G38").Interior.Color = xlNone
        .Range("G38").Value = ""
        
        .Range("G40").Interior.Color = xlNone
        .Range("G40").Value = ""
        .Range("G42").Interior.Color = xlNone
        .Range("G42").Value = ""
        .Range("G44").Interior.Color = xlNone
        .Range("G44").Value = ""
        .Range("G46").Interior.Color = xlNone
        .Range("G46").Value = ""
        .Range("G48").Interior.Color = xlNone
        .Range("G48").Value = ""
        
    End With

End Sub



Sub Save()

    Dim frm As Worksheet
    
    Dim database As Worksheet
    

    Dim iRow As Long

    Dim iSerial As Long
    
    Dim productLen As Integer
    
    Dim detailsLen As Integer
    
    Dim qtyLen As Integer
    
    Dim prjnameLen As Integer

   
    Set frm = ThisWorkbook.Sheets("Form")

    Set database = ThisWorkbook.Sheets("Database")
    
    
    Dim i As Integer
    
    Dim Products() As Variant
    If Trim(frm.Range("G10").Value) <> "" Then
        ReDim Preserve Products(0)
        'ReDim Preserve Products(UBound(Products) + 1)
        Products(0) = frm.Range("G10").Value
    End If
    
    If Trim(frm.Range("G10").Value) <> "" And Trim(frm.Range("G12").Value) <> "" Then
        ReDim Preserve Products(UBound(Products) + 1)
        Products(UBound(Products)) = frm.Range("G12").Value
    End If
    
    If Trim(frm.Range("G10").Value) <> "" And Trim(frm.Range("G12").Value) <> "" And Trim(frm.Range("G14").Value) <> "" Then
        ReDim Preserve Products(UBound(Products) + 1)
        Products(UBound(Products)) = frm.Range("G14").Value
    End If
    
    If Trim(frm.Range("G10").Value) <> "" And Trim(frm.Range("G12").Value) <> "" And Trim(frm.Range("G14").Value) <> "" And Trim(frm.Range("G16").Value) <> "" Then
        ReDim Preserve Products(UBound(Products) + 1)
        Products(UBound(Products)) = frm.Range("G16").Value
    End If
    
    If Trim(frm.Range("G10").Value) <> "" And Trim(frm.Range("G12").Value) <> "" And Trim(frm.Range("G14").Value) <> "" And Trim(frm.Range("G16").Value) <> "" And Trim(frm.Range("G18").Value) <> "" Then
        ReDim Preserve Products(UBound(Products) + 1)
        Products(UBound(Products)) = frm.Range("G18").Value
    End If
    
    
    Dim Details() As Variant
    If Trim(frm.Range("G20").Value) <> "" Then
        ReDim Preserve Details(0)
        'ReDim Preserve Details(UBound(Details) + 1)
        Details(0) = frm.Range("G20").Value
    End If
    
    If Trim(frm.Range("G20").Value) <> "" And Trim(frm.Range("G22").Value) <> "" Then
        ReDim Preserve Details(UBound(Details) + 1)
        Details(UBound(Details)) = frm.Range("G22").Value
    End If
    
    If Trim(frm.Range("G20").Value) <> "" And Trim(frm.Range("G22").Value) <> "" And Trim(frm.Range("G24").Value) <> "" Then
        ReDim Preserve Details(UBound(Details) + 1)
        Details(UBound(Details)) = frm.Range("G24").Value
    End If
    
    If Trim(frm.Range("G20").Value) <> "" And Trim(frm.Range("G22").Value) <> "" And Trim(frm.Range("G24").Value) <> "" And Trim(frm.Range("G26").Value) <> "" Then
        ReDim Preserve Details(UBound(Details) + 1)
        Details(UBound(Details)) = frm.Range("G26").Value
    End If
    
    If Trim(frm.Range("G20").Value) <> "" And Trim(frm.Range("G22").Value) <> "" And Trim(frm.Range("G24").Value) <> "" And Trim(frm.Range("G26").Value) <> "" And Trim(frm.Range("G28").Value) <> "" Then
        ReDim Preserve Details(UBound(Details) + 1)
        Details(UBound(Details)) = frm.Range("G28").Value
    End If
    
    
    Dim Qty() As Variant
    If Trim(frm.Range("G30").Value) <> "" Then
        ReDim Preserve Qty(0)
        'ReDim Preserve Qty(UBound(Qty) + 1)
        Qty(0) = frm.Range("G30").Value
    End If
    
    If Trim(frm.Range("G30").Value) <> "" And Trim(frm.Range("G32").Value) <> "" Then
        ReDim Preserve Qty(UBound(Qty) + 1)
        Qty(UBound(Qty)) = frm.Range("G32").Value
    End If
    
    If Trim(frm.Range("G30").Value) <> "" And Trim(frm.Range("G32").Value) <> "" And Trim(frm.Range("G34").Value) <> "" Then
        ReDim Preserve Qty(UBound(Qty) + 1)
        Qty(UBound(Qty)) = frm.Range("G34").Value
    End If
    
    If Trim(frm.Range("G30").Value) <> "" And Trim(frm.Range("G32").Value) <> "" And Trim(frm.Range("G34").Value) <> "" And Trim(frm.Range("G36").Value) <> "" Then
        ReDim Preserve Qty(UBound(Qty) + 1)
        Qty(UBound(Qty)) = frm.Range("G36").Value
    End If
    
    If Trim(frm.Range("G30").Value) <> "" And Trim(frm.Range("G32").Value) <> "" And Trim(frm.Range("G34").Value) <> "" And Trim(frm.Range("G36").Value) <> "" And Trim(frm.Range("G38").Value) <> "" Then
        ReDim Preserve Qty(UBound(Qty) + 1)
        Qty(UBound(Qty)) = frm.Range("G38").Value
    End If
    
    
    Dim PrjName() As Variant
    If Trim(frm.Range("G40").Value) <> "" Then
        ReDim Preserve PrjName(0)
        'ReDim Preserve PrjName(UBound(PrjName) + 1)
        PrjName(0) = frm.Range("G40").Value
    End If
    
    If Trim(frm.Range("G40").Value) <> "" And Trim(frm.Range("G42").Value) <> "" Then
        ReDim Preserve PrjName(UBound(PrjName) + 1)
        PrjName(UBound(PrjName)) = frm.Range("G42").Value
    End If
    
    If Trim(frm.Range("G40").Value) <> "" And Trim(frm.Range("G42").Value) <> "" And Trim(frm.Range("G44").Value) <> "" Then
        ReDim Preserve PrjName(UBound(PrjName) + 1)
        PrjName(UBound(PrjName)) = frm.Range("G44").Value
    End If
    
    If Trim(frm.Range("G40").Value) <> "" And Trim(frm.Range("G42").Value) <> "" And Trim(frm.Range("G44").Value) <> "" And Trim(frm.Range("G46").Value) <> "" Then
        ReDim Preserve PrjName(UBound(PrjName) + 1)
        PrjName(UBound(PrjName)) = frm.Range("G46").Value
    End If
    
    If Trim(frm.Range("G40").Value) <> "" And Trim(frm.Range("G42").Value) <> "" And Trim(frm.Range("G44").Value) <> "" And Trim(frm.Range("G46").Value) <> "" And Trim(frm.Range("G48").Value) <> "" Then
        ReDim Preserve PrjName(UBound(PrjName) + 1)
        PrjName(UBound(PrjName)) = frm.Range("G48").Value
    End If

   
    If Trim(frm.Range("L1").Value) = "" Then
        iRow = database.Range("A" & Application.Rows.Count).End(xlUp).Row + 1
        
        If iRow = 2 Then
            iSerial = 1
        Else
            iSerial = database.Cells(iRow - 1, 1).Value + 1
            
        End If
        
    Else
        iRow = frm.Range("K1").Value
        iSerial = frm.Range("L1").Value
        
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
            
                .Cells(iRow, 2).Offset(i, 0).Value = frm.Range("G6").Value
            
                .Cells(iRow, 3).Offset(i, 0).Value = frm.Range("G8").Value
            
                .Cells(iRow, 4).Offset(i, 0).Value = Products(i)
            
                .Cells(iRow, 5).Offset(i, 0).Value = Details(i)
           
                .Cells(iRow, 6).Offset(i, 0).Value = Qty(i)
            
                .Cells(iRow, 7).Offset(i, 0).Value = PrjName(i)
            
                .Cells(iRow, 8).Offset(i, 0).Value = Application.UserName
                
                
                orderText = "Order ID: " & (iSerial - 1) & vbNewLine & _
                            "Requested by: " & frm.Range("G6").Value & vbNewLine & _
                            "Requested Date: " & frm.Range("G8").Value & vbNewLine & _
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
    
    
    frm.Range("K1").Value = ""
    frm.Range("L1").Value = ""
        
End Sub



Sub Modify()

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

    Sheets("Form").Range("K1").Value = iRow
    Sheets("Form").Range("L1").Value = iSerial
    
    Sheets("Form").Range("G6").Value = Sheets("Database").Cells(iRow, 2).Value
    Sheets("Form").Range("G8").Value = Sheets("Database").Cells(iRow, 3).Value
    Sheets("Form").Range("G10").Value = Sheets("Database").Cells(iRow, 4).Value
    Sheets("Form").Range("G20").Value = Sheets("Database").Cells(iRow, 5).Value
    Sheets("Form").Range("G30").Value = Sheets("Database").Cells(iRow, 6).Value
    Sheets("Form").Range("G40").Value = Sheets("Database").Cells(iRow, 7).Value
    
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



