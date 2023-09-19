Sub AddContentToExistingDocument()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim ws As Worksheet
    Dim cell As Range
    Dim s18f As Boolean
    Dim s19f As Boolean
    Dim s305f As Boolean
    Dim s320f As Boolean
    Dim s321f As Boolean
    Dim s326f As Boolean
    Dim s331f As Boolean
    Dim s340f As Boolean
    Dim s347f As Boolean
    Dim s348f As Boolean
    Dim s350f As Boolean
    Dim s351f As Boolean
    Dim s405f As Boolean
    Dim s420f As Boolean
    Dim s430f As Boolean
    Dim s449f As Boolean
    Dim s395f As Boolean
    Dim s495f As Boolean
    Dim s515f As Boolean
    Dim s550f As Boolean
    Dim s551f As Boolean
    Dim s555f As Boolean
    Dim s556f As Boolean
    Dim s595f As Boolean
    Dim s630f As Boolean
    Dim s695f As Boolean
    Dim s700f As Boolean
    Dim s720f As Boolean
    Dim s730f As Boolean
    Dim s750f As Boolean
    Dim s770f As Boolean
    Dim s775f As Boolean
    Dim s780f As Boolean
    Dim s790f As Boolean
    Dim s795f As Boolean
    Dim s805f As Boolean
    Dim s810f As Boolean
    Dim s850f As Boolean
    Dim s857f As Boolean
    Dim s880f As Boolean


    ' Set the worksheet (replace "Budget" with your actual worksheet name)
    Set ws = ThisWorkbook.Worksheets("Budget")
   
    ' Try to get a reference to the existing Word application
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    On Error GoTo 0 ' Restore normal error handling

    ' Check if the Word application was found
    If wdApp Is Nothing Then
        ' If not found, create a new instance of Word
        Set wdApp = CreateObject("Word.Application")
    End If

    ' Try to get a reference to the existing document (change the document name and path)
    On Error Resume Next
    Set wdDoc = wdApp.Documents("WEI File Dividers - Company.docx")
    On Error GoTo 0 ' Restore normal error handling

    ' Check if the document was found
    If wdDoc Is Nothing Then
        ' If not found, open the existing document (change the document path)
        Set wdDoc = wdApp.Documents.Open("C:\Users\AvinashN\OneDrive - West-Evans Inc\Desktop\WEI File Dividers - Company.docx")
    End If
    
    ' Initialize the "found" variable to False
    s18f = False
    
    ' Loop through each cell in the range "C18:H18"
    For Each cell In ws.Range("C18:H18")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s18f = True
            wdDoc.Content.InsertAfter "18 - AFS PREPARATION" & vbFormFeed
            
            Exit For
        End If
    Next cell
    ' Initialize the "found" variable to False
    s19f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C27:H27")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s19f = True

       ' Insert "19 - IT CONTROLS"
        wdDoc.Content.InsertAfter "19 - IT CONTROLS" & vbFormFeed
        

            Exit For
    
        End If
    Next cell

        ' Initialize the "found" variable to False
    s305f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C28:H28")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s305f = True

            wdDoc.Content.InsertAfter "305 - PROPERTY,PLANT AND EQUIPMENT" & vbFormFeed
     
            Exit For
    
        End If
    Next cell
    
    ' Initialize the "found" variable to False
    s320f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C32:H32")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s320f = True

             wdDoc.Content.InsertAfter "320 - INTANGIBLE ASSETS" & vbFormFeed

            Exit For
        End If
    Next cell
    
        ' Initialize the "found" variable to False
    s321f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C33:H33")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s321f = True

            wdDoc.Content.InsertAfter "321 - GOODWILL" & vbFormFeed
            Exit For
        End If
    Next cell
    
            ' Initialize the "found" variable to False
    s326f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C34:H34")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s326f = True

            wdDoc.Content.InsertAfter "326 - INVESTMENT IN SUBSIDIARIES" & vbFormFeed
            Exit For
        End If
    Next cell
    
    ' Initialize the "found" variable to False
    s331f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C35:H35")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s331f = True

            wdDoc.Content.InsertAfter "331 - PREPAYMENTS" & vbFormFeed
            Exit For
        End If
    Next cell
    
    ' Initialize the "found" variable to False
    s340f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C36:H36")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s340f = True

            wdDoc.Content.InsertAfter "340 - LOAN RECEIVABLE" & vbFormFeed
            Exit For
        End If
    Next cell
    
        ' Initialize the "found" variable to False
    s347f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C37:H37")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s347f = True

            wdDoc.Content.InsertAfter "347 - LOANS TO GROUP COMPANIES" & vbFormFeed
            Exit For
        End If
    Next cell
    
        ' Initialize the "found" variable to False
    s348f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C38:H38")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s348f = True

            wdDoc.Content.InsertAfter "348 - LOANS WITH STAKEHOLDERS" & vbFormFeed
            Exit For
        End If
    Next cell
    
            ' Initialize the "found" variable to False
    s350f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C39:H39")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s350f = True

            wdDoc.Content.InsertAfter "350 - OTHER FINANCIAL ASSETS" & vbFormFeed
            Exit For
        End If
    Next cell
    
    ' Initialize the "found" variable to False
    s351f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C40:H40")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s351f = True

            wdDoc.Content.InsertAfter "351 - INVESTMENTS AT FAIR VALUE" & vbFormFeed
            Exit For
        End If
    Next cell
    
    ' Initialize the "found" variable to False
    s395f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C41:H41")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s351f = True

            wdDoc.Content.InsertAfter "395 - DEFFERRED TAX RECEIVABLE" & vbFormFeed
            Exit For
        End If
    Next cell
       
    
    
    
    
    ' Initialize the "found" variable to False
    s405f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C42:H42")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s405f = True

            wdDoc.Content.InsertAfter "405 - INVENTORIES" & vbFormFeed
            Exit For
        End If
    Next cell

    ' Initialize the "found" variable to False
    s420f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C44:H44")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s420f = True

            wdDoc.Content.InsertAfter "420 - CASH & CASH EQUIVALENTS" & vbFormFeed
            Exit For
        End If
    Next cell
    
        ' Initialize the "found" variable to False
    s430f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C45:H45")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s430f = True

            wdDoc.Content.InsertAfter "430 - TRADE AND OTHER RECEIVABLLES" & vbFormFeed
            Exit For
        End If
    Next cell
    
    ' Initialize the "found" variable to False
    s449f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C46:H46")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s449f = True

            wdDoc.Content.InsertAfter "449 - LOANS WITH EMPLOYEES" & vbFormFeed
            Exit For
        End If
    Next cell

    ' Initialize the "found" variable to False
    s495f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C47:H47")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s495f = True

            wdDoc.Content.InsertAfter "495 - CURRENT TAX RECEIVABLE" & vbFormFeed
            Exit For
        End If
    Next cell

    ' Initialize the "found" variable to False
    s515f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C48:H48")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s515f = True

            wdDoc.Content.InsertAfter "515 - PROVISIONS" & vbFormFeed
            Exit For
        End If
    Next cell
    
    
    ' Initialize the "found" variable to False
    s550f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C49:H49")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s550f = True

            wdDoc.Content.InsertAfter "550 - OTHER FINANCIAL ASSETS" & vbFormFeed
            Exit For
        End If
    Next cell
    
    
    ' Initialize the "found" variable to False
    s551f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C50:H50")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s551f = True

            wdDoc.Content.InsertAfter "551 - BORROWINGS" & vbFormFeed
            Exit For
        End If
    Next cell
    
    ' Initialize the "found" variable to False
    s555f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C51:H51")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s555f = True

            wdDoc.Content.InsertAfter "555 - LEASE LIABILITIES" & vbFormFeed
            Exit For
        End If
    Next cell
    
    ' Initialize the "found" variable to False
    s556f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C52:H52")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s556f = True

            wdDoc.Content.InsertAfter "556 - OPERATING LEASE LIABILITES" & vbFormFeed
            Exit For
        End If
    Next cell


    ' Initialize the "found" variable to False
    s595f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C53:H53")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s595f = True

            wdDoc.Content.InsertAfter "595 - DEFFERED TAX PAYABLE" & vbFormFeed
            Exit For
        End If
    Next cell
    
    ' Initialize the "found" variable to False
    s620f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C54:H54")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s620f = True

            wdDoc.Content.InsertAfter "620 - BANK OVERDRAFT" & vbFormFeed
            Exit For
        End If
    Next cell
    
        ' Initialize the "found" variable to False
    s630f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C55:H55")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s630f = True

            wdDoc.Content.InsertAfter "630 - TRADE AND OTHER PAYABLES" & vbFormFeed
            Exit For
        End If
    Next cell
    
    
    ' Initialize the "found" variable to False
    s695f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C56:H56")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s695f = True

            wdDoc.Content.InsertAfter "695 - CURRENT TAX PAYABLE" & vbFormFeed
            Exit For
        End If
    Next cell
    
    
    ' Initialize the "found" variable to False
    s700f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C57:H57")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s700f = True

            wdDoc.Content.InsertAfter "700 - REVENUE" & vbFormFeed
            Exit For
        End If
    Next cell


    ' Initialize the "found" variable to False
    s720f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C58:H58")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s720f = True

            wdDoc.Content.InsertAfter "720 - COST OF SALES" & vbFormFeed
            Exit For
        End If
    Next cell

    ' Initialize the "found" variable to False
    s730f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C59:H59")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s730f = True

            wdDoc.Content.InsertAfter "730 - OPERATING INCOME" & vbFormFeed
            Exit For
        End If
    Next cell
    
    ' Initialize the "found" variable to False
    s750f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C60:H60")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s750f = True

            wdDoc.Content.InsertAfter "750 - OPERATING EXPENSES" & vbFormFeed
            Exit For
        End If
    Next cell
    
    ' Initialize the "found" variable to False
    s770f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C61:H61")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s770f = True

            wdDoc.Content.InsertAfter "770 - INVESTMENT INCOME" & vbFormFeed
            Exit For
        End If
    Next cell
    
        ' Initialize the "found" variable to False
    s775f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C62:H62")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s775f = True

            wdDoc.Content.InsertAfter "775 - FINANCE COSTS" & vbFormFeed
            Exit For
        End If
    Next cell
    
        ' Initialize the "found" variable to False
    s780f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C63:H63")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s780f = True

            wdDoc.Content.InsertAfter "780-790 NON OPERATING GAINS/EXPENSES" & vbFormFeed
            Exit For
        End If
    Next cell
    
    
        ' Initialize the "found" variable to False
    s795f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C64:H64")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s795f = True

            wdDoc.Content.InsertAfter "795 - TAXATION" & vbFormFeed
            Exit For
        End If
    Next cell
    
    
            ' Initialize the "found" variable to False
    s805f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C65:H65")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s805f = True

            wdDoc.Content.InsertAfter "805 - CAPITAL" & vbFormFeed
            Exit For
        End If
    Next cell
    
    
    
    ' Initialize the "found" variable to False
    s810f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C66:H66")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s810f = True

            wdDoc.Content.InsertAfter "810-820 RETIANED EARNING/RESERVES " & vbFormFeed
            Exit For
        End If
    Next cell
    
    ' Initialize the "found" variable to False
    s850f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C67:H67")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s850f = True

            wdDoc.Content.InsertAfter "850 - RELATED PARTIES" & vbFormFeed
            Exit For
        End If
    Next cell
    
    ' Initialize the "found" variable to False
    s857f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C68:H68")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s857f = True

            wdDoc.Content.InsertAfter "857 - COMMITMENTS & CONTENGENCIES" & vbFormFeed
            Exit For
        End If
    Next cell
    
    
    ' Initialize the "found" variable to False
    s880f = False
    
    ' Loop through each cell in the range "C20:H20"
    For Each cell In ws.Range("C69:H69")
        ' Check if the cell's value is greater than 0.05
        If cell.Value > 0.05 Then
            ' Set the "found" variable to True and add content to the Word document
            s880f = True

            wdDoc.Content.InsertAfter "880 - CASHFLOWS" & vbFormFeed
            Exit For
        End If
    Next cell
    
    ' Save and close the document (if needed)
    wdDoc.Save
    'wdDoc.Close

    ' Show Word on the screen (if needed)
    wdApp.Visible = True
End Sub
