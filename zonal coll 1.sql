
Declare @dtDate  as smalldatetime
set @dtDate='14/MAr/2017'--'23/Feb/2017'
Select fltAmount,vchdescription,* From faVouchers Where tnyVouchertypeID=30 and dtDate=@dtDate Order by fltAmount

Select numLocationID,chvZoneNameEnglish,Sum(fltAmount) amt
From faVouchers 
Inner Join DB_Masters..GM_Zone M On M.numZoneID=faVouchers.numLocationID
Where tnyVoucherTypeID=10 and intInstrumentTypeID=1 And isnull(tnyCancelFlag,0)<>1 And intFinancialYearID=2016 And numLocationID<> 4016701
And dtDate=@dtDate
Group By numLocationID,chvZoneNameEnglish 
Order by amt
--Order by dtdate
Select fltAmount,* from faVouchers Where intFinancialyearID=2016 And tnyVouchertypeID=30 And dtDate=@dtDate And fltAmount  in (
Select Sum(fltAmount) amt
From faVouchers 
Inner Join DB_Masters..GM_Zone M On M.numZoneID=faVouchers.numLocationID
Where tnyVoucherTypeID=10 and intInstrumentTypeID=1 And isnull(tnyCancelFlag,0)<>1 And intFinancialYearID=2016 And numLocationID<> 4016701
And dtDate=@dtDate-- 
Group By numLocationID,chvZoneNameEnglish 
)Order by fltAmount

3065941.00

Select * from faVouchers Where intFinancialyearID=2016 And fltAmount=3065941.00

Select * from faVouchers Where intFinancialyearID=2016 And fltAmount=300000.00
.00



Select * from faVouchers Where intFinancialyearID=2016 And fltAmount=17479.00
.00

Select * from faVouchers Where intFinancialyearID=2016 And fltAmount=75463.00
Select * from faVouchers Where intFinancialyearID=2016 And fltAmount=100216.00
Select * from faVouchers Where intFinancialyearID=2016 And fltAmount=111832.00
Select * from faVouchers Where intFinancialyearID=2016 And fltAmount=269884.00

2961760,2961762,2961545  --attipra to cancel 857495.00



Select * From faVouchers Where intVoucherID in (2961762,2961545)
Select * From faVoucherChild Where intVoucherID in (2961762,2961545)



Select fltAmount,vchdescription,* From faVouchers Where tnyVouchertypeID=20 and dtDate='15/feb/2017' Order by fltAmount

Select * from faVouchers Where intFinancialyearID=2016 And fltAmount=99628.00

Select * from faVouchers Where intVoucherID in (2918248,2918781,2918548,2918610,2918318,2918881,2918381)



Select * from faVouchers Where intFinancialyearID=2016 And fltAmount=41810.00


2918837,2918870


Select * from faVouchers Where intFinancialyearID=2016 And fltAmount=204917.00

2752032,2756629,2755100,2751366,2755244





Select * From faVouchers Where intVoucherID=

Declare @VrID as numeric
set @VrID=2727437
Select * From faTransactions Where intVoucherID=@VrID
Select * From faTransactionChild Where intTransactionID in (Select intTransactionID From faTransactions Where intVoucherID=@VrID)
Select * From faVouchers Where intVoucherID in (@VrID)
Select * From faVoucherChild Where intVoucherID in (@VrID)

Select sum(fltAmount),tnyVouchertypeID from faVouchers Where dtDate='15/feb/2017' aND INTiNSTRUMENTtYPEid=1
Group by tnyVouchertypeID

Select fltAmount,tnyVouchertypeID,* from faVouchers Where dtDate='15/feb/2017' aND INTiNSTRUMENTtYPEid=1 and tnyVouchertypeID=20


Select * from faVouchers Where intVoucherID=2868360


Group by tnyVouchertypeID


125192.00	2704089	

Update faVouchers set fltAmount=125192.00 Where intVoucherID=2704089
Update faVoucherChild set fltAmount=125192.00 Where intVoucherID=2704089
Update faTransactionChild set fltAmount=125192.00 Where intTransactionID=2704049

Select * From faTransactions Where intVoucherID in (2704089,2704093)

Select * From faVouchers Where intVoucherID in (2868360)

Update faVouchers set tnyStatus=4 , tnyCancelFlag=1 Where intVoucherID=2704093



Select * from faVouchers Where intFinancialyearID=2016 And fltAmount=19573.00


Order by fltAmount

Update 