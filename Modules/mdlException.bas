Attribute VB_Name = "mdlException"
Option Explicit

Private Const EGYSEGAR_COL As Integer = 2
Private Const MENNYISEG_COL As Integer = 3
Private Const OSSZAR_COL As Integer = 4

Sub riportok()

On Error GoTo eh

    Dim utvonal As String
    utvonal = "C:\offlineTMP\youtube\kivetelkezeles-mint-egy-profi\zoldseges.xlsx"
    Dim wb As Workbook
    Set wb = Workbooks.Open(utvonal)
    
    Dim rg As Range
    Set rg = wb.Worksheets(1).Range("A1").CurrentRegion

    Debug.Print "Átlagos egységár: " & atlagEgysegar(rg) & " Ft"
    Debug.Print "Teljes készlet: " & teljesKeszlet(rg) & " kg"
    Debug.Print "Össz ár: " & osszAr(rg) & " Ft"

eh:
    wb.Close savechanges:=False
    Select Case Err.Number
        Case 0
            MsgBox "A riport elkészült!"
        Case vbObjectError + 1
            MsgBox "Ellenõrzés szükséges, mert:" & vbNewLine & Err.Description & vbNewLine & Err.Source
        Case Else
            MsgBox "Váratlan hiba történt!" & vbNewLine & Err.Description & vbNewLine & Err.Source
    End Select
End Sub

Function cellabolSzam(szam As Range, Optional ByRef szamlalo As Integer) As Double

On Error GoTo eh
    If IsNumeric(szam) Then
        cellabolSzam = CDbl(szam.Value)
        szamlalo = szamlalo + 1
    Else
'        cellabolSzam = 0
        Err.Raise vbObjectError + 1, "VBAProject", "A(z) " & szam.Address & " cellában nem szám szerepel!"
    End If
done:
    Exit Function
eh:
    Err.Raise Err.Number, Err.Source & vbNewLine & "cellabolSzam", Err.Description

End Function

Function atlagEgysegar(rg As Range) As Double

On Error GoTo eh
    
    Dim osszeg As Long
    Dim szamlalo As Integer
    
    Dim lv As Integer
    For lv = 2 To rg.Rows.Count
        osszeg = osszeg + cellabolSzam(rg(lv, EGYSEGAR_COL), szamlalo)
    Next lv
    
    atlagEgysegar = osszeg / szamlalo
        
done:
    Exit Function
eh:
    Err.Raise Err.Number, Err.Source & vbNewLine & "atlagEgysegar", Err.Description
End Function

Function teljesKeszlet(rg As Range) As Long

On Error GoTo eh

    Dim lv As Integer
    For lv = 2 To rg.Rows.Count
        teljesKeszlet = teljesKeszlet + rg(lv, MENNYISEG_COL)
        
done:
    Exit Function
eh:
    Err.Raise Err.Number, Err.Source & vbNewLine & "teljesKeszlet", Err.Description
    Next lv

End Function

Function osszAr(rg As Range) As Long

On Error GoTo eh
    
    Dim lv As Integer
    For lv = 2 To rg.Rows.Count
        osszAr = osszAr + cellabolSzam(rg(lv, OSSZAR_COL))
    Next lv
        
done:
    Exit Function
eh:
    Err.Raise Err.Number, Err.Source & vbNewLine & "osszAr", Err.Description
    
End Function

