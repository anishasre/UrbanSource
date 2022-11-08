Select dtdate ,intVoucherNo,fltAmount,vchDescription,* From faVouchers 
Where tnyVoucherTypeID=30 And intFinancialyearID=2016 and intInstrumentTypeID=1 and intTransactiontypeID=4001
And isNull(tnyReversed,0)<>1 Order By dtdate