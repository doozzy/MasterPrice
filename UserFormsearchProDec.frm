VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormsearchProDec 
   Caption         =   "Search based on product description"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9915
   OleObjectBlob   =   "UserFormsearchProDec.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormsearchProDec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Dim objWord As Object
    Dim subtotal As Double
    subtotal = 0
  
'   Start Word and create an object (late binding)
    Set objWord = CreateObject("Word.Application")
    'Set objDoc = objWord.Documents.Add
    
    
    '   Determine the filename
    'SaveAsName = Application.DefaultFilePath & _
           " \ " & Region & ".docx"
    SaveAsName = "C:\Users\Oldooz Dianat\Documents\SECOND.docx"
    MsgBox SaveAsName
    
    '   Add picture
    GraphImage = "C:\ProductMasterFile\ghs.png"
    
    
     
     With objWord
        '.Visible = True
        '.Activate
        .Documents.Add
        
        With .Selection
              .PageSetup.RightMargin = 30
            .PageSetup.LeftMargin = 30
            .PageSetup.TopMargin = 30
            .PageSetup.BottomMargin = 10
        
            Set objtable = .Tables.Add(Range:=objWord.Selection.Range, _
                    NumRows:=1, NumColumns:=2, _
                    DefaultTableBehavior:=wdWord9TableBehavior, _
                    AutoFitBehavior:=wdAutoFitContent)
                    
                objtable.Columns(1).PreferredWidth = 150
                objtable.Columns(2).PreferredWidth = 350
                objtable.Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
                Set wrdPic = .InlineShapes.AddPicture(Filename:=GraphImage, LinkToFile:=False, SaveWithDocument:=True)
                wrdPic.ScaleWidth = 50
                wrdPic.ScaleHeight = 50
            
                .moveRight
              
                .Font.Bold = True
                .TypeText Text:="General Healthcare Solution Pty Ltd" & Chr(11)
                .Font.Bold = False
                .TypeText Text:="ABN:36 161 023 418" & Chr(11) & "52 Gibbs St, Chatswood, NSW 2067, Australia" _
                & Chr(11) & "Ph: (61)2 9417 6566       Fax: (61)2 9417 5299" & vbCrLf & Chr(11)
                .Font.Bold = True
                .TypeText Text:="http://www.generalhealthcare.com.au"
                
        
        End With
        .Selection.moveDown
        With .Selection
            .ParagraphFormat.Alignment = 1
            .Font.Size = 14
            .Font.Bold = True
            .TypeText Text:="Quotation"
       
        End With
        
        With .Selection
           
            Set objtable = .Tables.Add(Range:=objWord.Selection.Range, _
                    NumRows:=1, NumColumns:=3, _
                    DefaultTableBehavior:=wdWord9TableBehavior, _
                    AutoFitBehavior:=wdAutoFitContent)
                    
                objtable.Columns(1).PreferredWidth = 230
                objtable.Columns(2).PreferredWidth = 40
                objtable.Columns(3).PreferredWidth = 230
                objtable.Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
                objtable.Cell(1, 1).Borders.Enable = True
                             
               
                .Font.Bold = True
                .Font.Size = 12
                .TypeText Text:=Chr(11)
                .TypeText Text:="ATTENTION TO:" & Chr(11)
                .moveRight
                
                .moveRight
                objtable.Cell(1, 3).Borders.Enable = True
                objtable.Cell(1, 3).Range.ParagraphFormat.Alignment = 0
                .Font.Size = 12
                .Font.Bold = True
                .TypeText Text:=Chr(11)
                .TypeText Text:="Quotation No." & Chr(11)
                .TypeText Text:="Date:" & Chr(9) & Chr(9)
                .Font.Bold = False
                .TypeText Text:=Chr(9) & Date & Chr(11)
                .Font.Bold = True
                .TypeText Text:="Payment Terms: " & Chr(11)
                .TypeText Text:="Ship via:" & Chr(9) & Chr(9)
                 If OptionFreight Then
                 .Font.Size = 12
                 .Font.Bold = False
                 .TypeText Text:="Freight" & Chr(11)
                 End If
                If Optionair Then
                .Font.Size = 12
                .Font.Bold = False
                .TypeText Text:="Air" & Chr(11)
                End If
                .Font.Bold = True
                .TypeText Text:="Salesperson:" & Chr(11)
        
        End With
        .Selection.moveDown
        
          With .Selection
            .ParagraphFormat.Alignment = 1
            .Font.Bold = False
            .TypeText Text:=String(75, "_")
       
          End With
          
         With .Selection
         defaultrows = 8
            Set objtable = .Tables.Add(Range:=objWord.Selection.Range, _
                    NumRows:=2 + defaultrows, NumColumns:=6, _
                    DefaultTableBehavior:=wdWord9TableBehavior, _
                    AutoFitBehavior:=wdAutoFitContent)
                    objtable.Borders.Enable = True
                 objtable.Rows(1).HeadingFormat = True
                'objtable.Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
           
                  .Font.Size = 12
                .Font.Bold = True
                 '.Borders(xlEdgeTop).Weight = xlThin
                 'objtable.Borders.OutsideLineStyle = True
                 objtable.Borders.InsideLineStyle = wdLineStyleSingle
                'objTable.Rows(1).Cells.Borders.Item(ppBorderTop).Visible = msoTrue
               'objTable.Borders.OutsideLineStyle = wdLineStyleSingle
               'objTable.Borders.OutsideLineWidth = wdLineWidth075pt
               ' objTable.Borders.OutsideColor = wdBlack
               
                objtable.Rows(10).Borders.Enable = True
                objtable.Columns.Borders.Enable = True
                objtable.Rows.Borders.Enable = False
                
                objtable.Rows(1).Borders.Enable = True
                
                
               ' objtable.Condition(wdLastRow).Borders.Enable = True
                'objtable.Borders(wdBorderBottom).InsideLineStyle = wdLineStyleSingle
 'objTable.Borders.OutsideLineStyle = wdLineStyleDouble
                'objTable.Borders.InsideLineStyle = wdLineStyleNone
                objtable.Columns(1).PreferredWidth = 50
                objtable.Columns(2).PreferredWidth = 250
               
                objtable.Columns(3).PreferredWidth = 40
                objtable.Columns(4).PreferredWidth = 70
                objtable.Columns(5).PreferredWidth = 50
                objtable.Columns(6).PreferredWidth = 90
                
                
                 objtable.Cell(1, 1).Range.Text = "Item No."
                  objtable.Cell(1, 2).Range.Font.Size = 12
                objtable.Cell(1, 2).Range.Font.Bold = True
                objtable.Cell(1, 2).Range.Text = "Product Description"
                 objtable.Cell(1, 3).Range.Font.Size = 12
                objtable.Cell(1, 3).Range.Font.Bold = True
                objtable.Cell(1, 3).Range.Text = "Box Qty"
        
                 objtable.Cell(1, 4).Range.Font.Size = 12
                objtable.Cell(1, 4).Range.Font.Bold = True
                objtable.Cell(1, 4).Range.Text = "Unit Price"
                 objtable.Cell(1, 5).Range.Font.Size = 12
                objtable.Cell(1, 5).Range.Font.Bold = True
                objtable.Cell(1, 5).Range.Text = "UOM Unit"
                 objtable.Cell(1, 6).Range.Font.Size = 12
                objtable.Cell(1, 6).Range.Font.Bold = True
                objtable.Cell(1, 6).Range.Text = "Box Price"
               
                
                j = 1
                For i = 0 To ListBoxSearchSpecificColumn.ListCount - 1
                    If ListBoxSearchSpecificColumn.Selected(i) = True Then
                    'ListBox2.AddItem ListBox1.List(i)
                    objtable.Cell(1 + j, 1).Range.Font.Size = 10
                    objtable.Cell(1 + j, 1).Range.Text = ListBoxSearchSpecificColumn.List(i, 0)
                    objtable.Cell(1 + j, 2).Range.Font.Size = 10
                    objtable.Cell(1 + j, 2).Range.Text = ListBoxSearchSpecificColumn.List(i, 1)
                    objtable.Cell(1 + j, 3).Range.Font.Size = 10
                    objtable.Cell(1 + j, 3).Range.Text = ListBoxSearchSpecificColumn.List(i, 2)
                    objtable.Cell(1 + j, 4).Range.Font.Size = 10
                    objtable.Cell(1 + j, 4).Range.Text = ListBoxSearchSpecificColumn.List(i, 3)
                    objtable.Cell(1 + j, 5).Range.Font.Size = 10
                    objtable.Cell(1 + j, 5).Range.Text = ListBoxSearchSpecificColumn.List(i, 4)
                    objtable.Cell(1 + j, 6).Range.Font.Size = 10
                    objtable.Cell(1 + j, 6).Range.Text = ListBoxSearchSpecificColumn.List(i, 5)
                    If Not IsNumeric(ListBoxSearchSpecificColumn.List(i, 5)) Then
                    ListBoxSearchSpecificColumn.List(i, 5) = 0
                    End If
                    
                    subtotal = subtotal + ListBoxSearchSpecificColumn.List(i, 5)
                  
                    j = j + 1
                    End If
                Next i
                
                'objTable.Borders.OutsideLineStyle = wdLineStyleSingle
                'objTable.Borders(xlEdgeBottom).LineStyle = xlContinuous
                
                For Each wdCell In objtable.Columns(1).Cells
                    .moveDown
                Next wdCell
                .moveDown
        End With
      
        With .Selection
            .ParagraphFormat.Alignment = 0
            .Font.Size = 10
            .Font.Bold = True
            .TypeText Text:=Chr(11)
            .TypeText Text:="Quotation valid 90 days"
       
        End With
        
        With .Selection
            Set objtable = .Tables.Add(Range:=objWord.Selection.Range, _
                    NumRows:=2, NumColumns:=3, _
                    DefaultTableBehavior:=wdWord9TableBehavior, _
                    AutoFitBehavior:=wdAutoFitContent)
                    objtable.Borders.Enable = True
                 'objtable.Rows(1).HeadingFormat = True
                'objtable.Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
                     objtable.Columns(1).PreferredWidth = 300
                objtable.Columns(2).PreferredWidth = 100
                objtable.Columns(3).PreferredWidth = 100
                objtable.Columns(1).Borders.Enable = False
                objtable.Columns(2).Borders.Enable = True
                objtable.Columns(3).Borders.Enable = True
                .moveRight
                
                  .Font.Size = 12
                .Font.Bold = True
                .TypeText Text:="Subtotal" & Chr(11)
                .Font.Bold = False
                .TypeText Text:="Freight Charge"
                .moveRight
                .TypeText Text:="$" & subtotal & Chr(11)
                .TypeText Text:="$ 0 "
               
                
                .moveDown
                .moveLeft
                  .Font.Size = 12
                .Font.Bold = True
                .TypeText Text:="TOTAL INVOICE"
                .moveRight
                .TypeText Text:="$" & subtotal & " FReCha"
                 
                .moveDown
        End With
        .Selection.moveDown
         With .Selection
            .ParagraphFormat.Alignment = 2
            .Font.Size = 10
            .Font.Bold = False
            .TypeText Text:=Chr(11)
            .TypeText Text:="Exclude G.S.T"
       
        End With
        
        .ActiveDocument.SaveAs Filename:=SaveAsName
    End With

   
 
    
    '   Kill the object
    objWord.Quit
    Set objWord = Nothing
'   Reset status bar
    Application.StatusBar = ""
    MsgBox Records & " memos were created and saved in " & _
      Application.DefaultFilePath
    

End Sub



Private Sub SaveAsPDF_Click()
'   Creates memos in word using Automation
    Dim WordApp As Object
    Dim Data As Range, message As String
    Dim Records As Integer, i As Integer
    Dim Region As String, SalesAmt As String, SalesNum As String
    Dim SaveAsName As String
'   Start Word and create an object (late binding)
    Set WordApp = CreateObject("Word.Application")
    
'   Add picture
    GraphImage = "C:\ProductMasterFile\ghs.png"
    
'   Information from worksheet
    Set Data = Sheets("ACTIVE 2011").Range("A4")
    'message = Sheets("ACTIVE 2011").Range("Message")
    
'   Cycle through all records in Sheet1
    Records = Application.CountA(Sheets("ACTIVE 2011").Range("A4:A6"))
    
'   Update status bar progress message
    'Application.StatusBar = "Processing Record " & i
'   Assign current data to variables
    Region = Data.Cells(1, 2).Value
    SalesNum = Data.Cells(1, 1).Value
    SalesAmt = Format(Data.Cells(1, 4).Value, "#,000")
        
'   Determine the filename
    'SaveAsName = Application.DefaultFilePath & _
           " \ " & Region & ".docx"
    SaveAsName = "C:\Users\Oldooz Dianat\Documents\first.docx"
    MsgBox SaveAsName
 
     

    Dim intNoOfRows As Integer
    Dim intNoOfColumns As Integer
    
    intNoOfRows = 1

    intNoOfColumns = 2


    'For i = 1 To 1

       ' For j = 1 To intNoOfColumns

            'objTable.Cell(i, j).Range.Text = "Sumit_" & i & j

        'Next

    'Next

    
     
'   Send commands to Word
    With WordApp
        .Documents.Add
        With .Selection
            .PageSetup.RightMargin = 30
            .PageSetup.LeftMargin = 30
            .PageSetup.TopMargin = 30
            .PageSetup.BottomMargin = 10
                                  
            Set wrdPic = .InlineShapes.AddPicture(Filename:=GraphImage, LinkToFile:=False, SaveWithDocument:=True)
            wrdPic.ScaleWidth = 50
            wrdPic.ScaleHeight = 50
         
            .Font.Size = 12
            .Font.Bold = True
            .ParagraphFormat.Alignment = 1
            .TypeText Text:="General Healthcare Solution Pty Ltd" & Chr(11)
            .Font.Bold = False
            .TypeText Text:=vbTab & "ABN:36 161 023 418" & Chr(11) & "52 Gibbs St, Chatswood, NSW 2067, Australia" _
            & Chr(11) & "Ph: (61)2 9417 6566 Fax: (61)2 9417 5299" & vbCrLf
            .ParagraphFormat.Alignment = 1
            .Font.Size = 14
            .Font.Bold = True
            .TypeText Text:="Quotation"
            .TypeParagraph
            .TypeParagraph
            .Font.Size = 12
            .ParagraphFormat.Alignment = 0
            .Font.Bold = False
            .TypeText Text:="Date:" & vbTab & _
            Format(Date, "mmmm d, yyyy")
            .TypeParagraph
            .TypeText Text:="To:" & vbTab & Region & _
             " Manager"
             .TypeParagraph
            .TypeText Text:="From:" & vbTab & _
               Application.UserName
            .TypeParagraph
            .TypeParagraph
            '.TypeText message
            .TypeParagraph
            .TypeParagraph
            .TypeText Text:="Units Sold:" & vbTab & _
             SalesNum
            .TypeParagraph
            .TypeText Text:="Amount:" & vbTab & _
             Format(SalesAmt, "$#,##0")
        End With
        .ActiveDocument.SaveAs Filename:=SaveAsName
    End With
    
'   Kill the object
    WordApp.Quit
    Set WordApp = Nothing
'   Reset status bar
    Application.StatusBar = ""
    MsgBox Records & " memos were created and saved in " & _
      Application.DefaultFilePath
End Sub


Private Sub searchbutton_Click()
'
'Search a product description
'
        ListBoxSearchSpecificColumn.Clear

        If Not searchInput.Value = "" Then
            GoTo productValid
        Else
            MsgBox " That product not exist. Please try again.", Title:="GHS"
        End If
        Exit Sub
    
    
productValid:
    Dim j As Integer
    Dim aCell As String
    Dim flag As Boolean
    flag = False
    
    
    j = 0
    aCell = searchInput.Text
    
    ListBoxSearchSpecificColumn.AddItem
    ListBoxSearchSpecificColumn.FontBold = True
   ListBoxSearchSpecificColumn.List(0, 0) = "Item No."
   ListBoxSearchSpecificColumn.List(0, 1) = "Production Description"
   ListBoxSearchSpecificColumn.List(0, 2) = "Box Qty"
   ListBoxSearchSpecificColumn.List(0, 3) = "Unit Price"
   ListBoxSearchSpecificColumn.List(0, 4) = "UOM Unit"
   ListBoxSearchSpecificColumn.List(0, 5) = "Box Price"
   
   
   ListBoxSearchSpecificColumn.FontBold = False
   j = 1
    For i = 3 To 128
        If InStr(Sheets("ACTIVE 2011").Range("B" & i).Value, aCell) > 0 Then
            'ListBoxSearchSpecificColumn.RowSource = "'ACTIVE 2011'!" & Sheets("ACTIVE 2011").Range("A" & i & ":" & "B" & i & ";"& "A" & i & ":" & "B").Address
            ListBoxSearchSpecificColumn.AddItem
            ListBoxSearchSpecificColumn.List(j, 0) = Sheets("ACTIVE 2011").Range("B" & i)
            ListBoxSearchSpecificColumn.List(j, 1) = Sheets("ACTIVE 2011").Range("C" & i)
            ListBoxSearchSpecificColumn.List(j, 2) = Sheets("ACTIVE 2011").Range("D" & i)
            ListBoxSearchSpecificColumn.List(j, 3) = Sheets("ACTIVE 2011").Range("Q" & i)
            ListBoxSearchSpecificColumn.List(j, 4) = Sheets("ACTIVE 2011").Range("E" & i)
            If IsNumeric(ListBoxSearchSpecificColumn.List(j, 2)) And IsNumeric(ListBoxSearchSpecificColumn.List(j, 3)) Then
            ListBoxSearchSpecificColumn.List(j, 5) = ListBoxSearchSpecificColumn.List(j, 2) * ListBoxSearchSpecificColumn.List(j, 3)
            Else
             ListBoxSearchSpecificColumn.List(j, 5) = "Box Qty or Unit Price is not available."
           End If
            
            j = j + 1
        'Else
           ' MsgBox " That product not exist. Please try again.", Title:="GHS"
            'Exit For
            flag = True
        End If
    Next i
    If flag = False Then
        MsgBox " That product not exist. Please try again.", Title:="GHS"
    End If
    
End Sub
