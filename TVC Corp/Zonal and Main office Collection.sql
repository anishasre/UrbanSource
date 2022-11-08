
Select dtDate,numLocationID,chvZoneNameEnglish,Sum(fltAmount) 
From faVouchers 
Inner Join DB_Masters..GM_Zone M On M.numZoneID=faVouchers.numLocationID
Where tnyVoucherTypeID=10 and intInstrumentTypeID=1 And isnull(tnyCancelFlag,0)<>1 And intFinancialYearID=2016 And numLocationID<> 4016701
Group By numLocationID,dtDate,chvZoneNameEnglish 
--Order by dtdate

Union All

Select dtDate,numLocationID,chvZoneNameEnglish,Sum(fltAmount) 
From faVouchers 
Inner Join faCounters On faCounters.intCounterID=faVouchers.intCounterID
Inner Join DB_Masters..GM_Zone M On M.numZoneID=faVouchers.numLocationID
Where tnyVoucherTypeID=10 and intInstrumentTypeID=1 And isnull(tnyCancelFlag,0)<>1 And intFinancialYearID=2016 And numLocationID= 4016701
And intSectionID=99
Group By numLocationID,dtDate,chvZoneNameEnglish

--Order by dtdate

