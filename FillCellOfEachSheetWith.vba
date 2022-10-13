' Remplit toutes les cellules B1 de toutes les feuilles avec le mois et l'année du document,
' en majuscule
' Par exemple: NOVEMBRE 2022 pour la feuille de planning de novembre 2022
' ---
Sub fillSheetsDates(pYear As String)
    Dim count As Byte
    Dim dateFormula As String
    
    count = 1
    For Each Sheet In Worksheets
        If Sheet.Name <> "Styles" Then
            dateFormula = "=UPPER(TEXT(""01/" & count & "/" & pYear & """, ""mmmm aaaa""))"
            Sheet.Range("B1").Formula = dateFormula
            count = count + 1
        End If
    Next Sheet
End Sub
