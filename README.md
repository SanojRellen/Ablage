Sub MyMacro()
    ' Löscht die Spalten C, D und G
    Columns("C:C").Delete
    Columns("D:D").Delete
    Columns("G:G").Delete
    
    ' Ersetzt Punkt durch Komma in C25
    Range("C25").Value = Replace(Range("C25").Value, ".", ",")
    
    ' Ersetzt Komma durch Punkt in C21
    Range("C21").Value = Replace(Range("C21").Value, ",", ".")
    
    ' Löscht die Zeilen 36, 20 und 8
    Rows(36).Delete
    Rows(20).Delete
    Rows(8).Delete
End Sub
