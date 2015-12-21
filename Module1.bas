Attribute VB_Name = "Module1"
Sub exchangeRate()
'
' ExchangeRate Macro
'

'
    Dim exchangeRate As Double
    
    exchangeRate = InputBox("Please enter the current exchange rate.")
    
    
    'If IsNull(exchangeRate) = 0 Then Exit Sub
    
    ThisWorkbook.Sheets("ACTIVE 2011").Range("O3:O182").FormulaR1C1 = "=RC[-2]*" & exchangeRate & ""
    
    ThisWorkbook.Sheets("ACTIVE 2011").Range("L3:L182").FormulaR1C1 = "=IFERROR(RC[-1]*" & exchangeRate & ", ""Not available"")"
    
End Sub

Sub searchProductDes()
    UserFormsearchProDec.Show vbModeless
    
End Sub

