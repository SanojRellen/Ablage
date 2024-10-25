=WORKDAY(EOMONTH(A1, 5), 1 + (2 - WEEKDAY(EOMONTH(A1, 5) + 1, 2)) % 7)

=Y32 + CHOOSE(WEEKDAY(Y32, 2), 0, 6, 5, 4, 3, 2, 1)



=INDEX(Sheet1!X:X, MATCH(A1, Sheet1!B:B, 0) + 4)

