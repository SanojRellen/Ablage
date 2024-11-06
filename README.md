
=IF(WEEKDAY(TODAY())<=5, TODAY() - (ROW()-1) - INT((ROW()-1)/5)*2, TODAY() - ROW() - INT(ROW()/5)*2)
