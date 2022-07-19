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

    Debug.Print "�tlagos egys�g�r: " & atlagEgysegar(rg) & " Ft"
    Debug.Print "Teljes k�szlet: " & teljesKeszlet(rg) & " kg"
    Debug.Print "�ssz �r: " & osszAr(rg) & " Ft"

eh:
    wb.Close savechanges:=False
    Select Case Err.Number
        Case 0
            MsgBox "A riport elk�sz�lt!"
        Case vbObjectError + 1
            MsgBox "Ellen�rz�s sz�ks�ges, mert:" & vbNewLine & Err.Description & vbNewLine & Err.Source
        Case Else
            MsgBox "V�ratlan hiba t�rt�nt!" & vbNewLine & Err.Description & vbNewLine & Err.Source
    End Select
End Sub

Function cellabolSzam(szam As Range, Optional ByRef szamlalo As Integer) As Double

On Error GoTo eh
    If IsNumeric(szam) Then
        cellabolSzam = CDbl(szam.Value)
        szamlalo = szamlalo + 1
    Else
'        cellabolSzam = 0
        Err.Raise vbObjectError + 1, "VBAProject", "A(z) " & szam.Address & " cell�ban nem sz�m szerepel!"
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

