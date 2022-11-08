
Select Replace(vchDescription,' ',''),vchDescription + ' '+ convert(varchar(15),intVoucherID),intVoucherID,dtDate,fltAmount,* From faVouchers Where tnyVoucherTypeID=30 and intTransactionTypeID=4001 and intInstrumentTypeID=1 And intFinancialyearID=2016 and
intVoucherId not in (Select intVoucherID From faVouchers Where tnyVoucherTypeID=30 and intTransactionTypeID=4001 and intInstrumentTypeID=1 and vchDescription like '%Jsk%')
--And intVoucherId in (Select intVoucherID From faVouchers Where tnyVoucherTypeID=30 and intTransactionTypeID=4001 and intInstrumentTypeID=1 and vchDescription like '%%')
 And dtDate between '28/Feb/2017'  and '28/Mar/2017'

And vchDescription like '%kaz%'



 
Select * From faVouchers Where intvoucherID in (2727304)

select numLocationID,sum(fltAmount) From faVouchers Where intFinancialyearID=2016 and dtDate='1/Apr/2016' and numLocationID<>4016701
Group By numLocationID
Select * From faVouchers Where intVoucherID in (2643030,2643197,2643552,2643669,2643739,2668246,2668299,2668386




Select * From faVouchers Where intVoucherID in (2643026,2643193,2643420,2643548,2643664,2643736,2667815,2668297,2668384,2699989
,2700119,2700261,2700477,2700673,2702670)
4016701


Select fltAmount,vchDescription,intVoucherID,dtDate,* From faVouchers Where tnyVoucherTypeID=30 and intTransactionTypeID=4001 and intInstrumentTypeID=1 And intFinancialyearID=2016




Update  faVouchers set dtDate='1/Apr/2016' Where intVoucherID in (2643026,2643193,2643420,2643548,2643664,2643736,2667815,2668297,2668384)
Update  faTransactions set dtTransactionDate='1/Apr/2016' Where intVoucherID in (2643026,2643193,2643420,2643548,2643664,2643736,2667815,2668297,2668384)



