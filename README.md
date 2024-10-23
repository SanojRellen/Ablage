=MID(A1, FIND(" ", A1, FIND(" ", A1, FIND(" ", A1) + 1) + 1) + 1, 1)


=INDEX(N8:N10, MATCH(1, (L8:L10=D8)*(O8:O10=G8), 0))



=AND(A1>=TODAY(), A1<=TODAY()+60)


=DATEVALUE(TEXT(BDP("AAPL US Equity", "EARN_ANN_DT"),"mm/dd/yyyy"))


=DATEVALUE(BDP("AAPL US Equity", "EARN_ANN_DT"))
