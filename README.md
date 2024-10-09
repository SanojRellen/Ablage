Range("A1").Value = Replace(Replace(Format(Range("A1").Value, "#,##0.00"), ".", "X"), ",", ".")
Range("A1").Value = Replace(Range("A1").Value, "X", ",")
