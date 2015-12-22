Attribute VB_Name = "Module1"
Sub exchangeRate()
'
' ExchangeRate Macro
'

'
    Dim exchangeRate As String
    
    exchangeRate = InputBox("Please enter the current exchange rate.")
   
    If Not IsNumeric(exchangeRate) Then
     MsgBox "You should enter a number"
     Exit Sub
     End If
    
    
    ThisWorkbook.Sheets("ACTIVE 2011").Range("O3:O182").FormulaR1C1 = "=RC[-2]*" & CDbl(exchangeRate) & ""
    
    ThisWorkbook.Sheets("ACTIVE 2011").Range("L3:L182").FormulaR1C1 = "=IFERROR(RC[-1]*" & CDbl(exchangeRate) & ", ""Not available"")"
    
End Sub

Sub searchProductDes()
    UserFormsearchProDec.Show vbModeless
    
End Sub

