Attribute VB_Name = "mdlException"
Option Explicit

Private Const EGYSEGAR_COL As Integer = 2
Private Const MENNYISEG_COL As Integer = 3
Private Const OSSZAR_COL As Integer = 4

Sub riportok()

On Error GoTo eh
    
    Dim utvonal As String
    utvonal = "" 'fajl helye
    Dim wb As Workbook
    Set wb = Workbooks.Open(utvonal)
    
    Dim rg As Range
    Set rg = wb.Worksheets(1).Range("A1").CurrentRegion

    Debug.Print "Átlagos egységár: " & atlagEgysegar(rg) & " Ft"
    Debug.Print "Teljes készlet: " & teljesKeszlet(rg) & " kg"
    Debug.Print "Össz ár: " & osszAr(rg) & " Ft"

eh:
    wb.Close SaveChanges:=False
    Select Case Err.Number
        Case 0
            MsgBox "A riport elkészült!"
        Case vbObjectError + 555
            MsgBox "Oops: " & vbNewLine & Err.Description & vbNewLine & Err.Source
        Case Else
            MsgBox "Hiba történt: " & vbNewLine & Err.Description & vbNewLine & Err.Source
    End Select
End Sub

Function cellabolSzam(ertek As Range) As Double

On Error GoTo eh

    If IsNumeric(ertek.Value) Then
        cellabolSzam = CDbl(ertek.Value)
    Else
        Err.Raise _
            vbObjectError + 555, _
            "VBAProject", _
            "A(z) " & ertek.Address & " cella értéke nem szám!"
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
        osszeg = osszeg + cellabolSzam(rg(lv, EGYSEGAR_COL))
        szamlalo = szamlalo + 1
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
        teljesKeszlet = teljesKeszlet + cellabolSzam(rg(lv, MENNYISEG_COL))
    Next lv

done:
    Exit Function
eh:
    Err.Raise Err.Number, Err.Source & vbNewLine & "teljesKeszlet", Err.Description
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

