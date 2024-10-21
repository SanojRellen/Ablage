=MID(A1, FIND(" ", A1, FIND(" ", A1, FIND(" ", A1) + 1) + 1) + 1, 1)


=INDEX(N8:N10, MATCH(1, (L8:L10=D8)*(O8:O10=G8), 0))

