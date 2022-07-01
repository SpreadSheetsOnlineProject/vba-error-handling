Attribute VB_Name = "mdlException"
Option Explicit

Sub riportok()
On Error GoTo eh
    
    Dim utvonal As String
    utvonal = "" 'ide johet a fajl elelresi utja
    Dim wb As Workbook
    Set wb = Workbooks.Open(utvonal)
    
    Dim rg As Range
    Set rg = wb.Worksheets(1).Range("A1").CurrentRegion

    Debug.Print "�tlagos egys�g�r: " & atlagEgysegar(rg) & " Ft"
    Debug.Print "Teljes k�szlet: " & teljesKeszlet(rg) & " kg"
    Debug.Print "�ssz �r: " & osszAr(rg) & " Ft"

eh:
    wb.Close SaveChanges:=False
    
    Select Case Err.Number
        Case 0
            Debug.Print "A riport elk�sz�lt!"
        Case vbObjectError + 555
            Debug.Print "Csak sz�mokkal v�gezhet�ek el a m�veletek!" & vbNewLine & "Hiba oka: " & Err.Description & vbNewLine & "Hiba forr�sa: " & Err.Source
        Case Else
            Debug.Print "V�ratlan hiba t�rt�nt!" & vbNewLine & "Hiba oka: " & Err.Description & vbNewLine & "Hiba forr�sa: " & Err.Source
    End Select

End Sub

Function cellabolSzam(ByRef szam As Range) As Double
On Error GoTo eh

    If IsNumeric(szam.Value2) Then
        cellabolSzam = CDbl(szam)
    Else
        Err.Raise _
            vbObjectError + 555, _
            "VBAProject", _
            "A(z) " & szam.Address & " cella �rt�ke nem sz�m!"
    End If
    
eh:
    Select Case Err.Number
        Case 0: Exit Function
        Case Else
            Err.Raise Err.Number, Err.Source & vbNewLine & "cellabolSzam", Err.Description
    End Select
End Function

Function atlagEgysegar(rg As Range) As Double
On Error GoTo eh

    Const EGYSEGAR_COL As Integer = 2
    
    Dim osszeg As Long
    
    Dim lv As Integer
    For lv = 2 To rg.Rows.Count
        osszeg = osszeg + cellabolSzam(rg(lv, EGYSEGAR_COL))
    Next lv
    
    atlagEgysegar = osszeg / (rg.Rows.Count - 1)
    
eh:
    Select Case Err.Number
        Case 0: Exit Function
        Case Else
            Err.Raise Err.Number, Err.Source & vbNewLine & "atlagEgysegar", Err.Description
    End Select
    
End Function

Function teljesKeszlet(rg As Range) As Long
On Error GoTo eh
    
    Const MENNYISEG_COL As Integer = 3

    Dim lv As Integer
    For lv = 2 To rg.Rows.Count
        teljesKeszlet = teljesKeszlet + cellabolSzam(rg(lv, MENNYISEG_COL))
    Next lv

eh:
    Select Case Err.Number
        Case 0: Exit Function
        Case Else
            Err.Raise Err.Number, Err.Source & vbNewLine & "teljesKeszlet", Err.Description
    End Select
End Function

Function osszAr(rg As Range) As Long
On Error GoTo eh
    
    Const OSSZAR_COL As Integer = 4
    
    Dim lv As Integer
    For lv = 2 To rg.Rows.Count
        osszAr = osszAr + cellabolSzam(rg(lv, OSSZAR_COL))
    Next lv
    
eh:
    Select Case Err.Number
        Case 0: Exit Function
        Case Else
            Err.Raise Err.Number, Err.Source & vbNewLine & "osszAr", Err.Description
    End Select
End Function

