Select dtDate,numLocationID,chvZoneNameEnglish,Sum(fltAmount) 
From faVouchers 
Inner Join GM_Zone On GM_Zone.numZoneID=faVouchers.numLocationID
Where tnyVoucherTypeID=10 and intInstrumentTypeID=1 And isnull(tnyCancelFlag,0)<>1
Group By numLocationID,dtDate,chvZoneNameEnglish
Order By dtdate



Select * From  faVouchers


Select * From GM_Zone