Delete From faVoucherChild Where intVoucherID in (Delete From faVouchers Where intFinancialYearID<2018)
Delete From faVoucherSub Where intVoucherID in (Delete From faVouchers Where intFinancialYearID<2018)
Delete From faVoucherAddress Where intVoucherID in (Delete From faVouchers Where intFinancialYearID<2018)
Delete From faVouchers Where intFinancialYearID<2018


Select *  From faVouchers Where intFinancialYearID= 2011 and intTransactionTypeID<>3000

Select * From faVouchers Where intTransactionTypeID=3000

--  2010   2011-03-31 00:00:00

--2011
Select *  From faVouchers Where intFinancialYearID= 2011 and intTransactionTypeID<>3000
Select * From faVoucherChild Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2011 and intTransactionTypeID<>3000)
Select * From faVoucherSub Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2011 and intTransactionTypeID<>3000)
Select * From faVoucherAddress Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2011 and intTransactionTypeID<>3000)
Select * From faTransactions Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2011 and intTransactionTypeID<>3000)
Select * From faTransactionChild Where intTransactionId in (Select intTransactionID From faTransactions Where intVoucherID in 
(Select intVoucherID  From faVouchers Where intFinancialYearID= 2011 and intTransactionTypeID<>3000))



Delete From faVoucherChild Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2011 and intTransactionTypeID<>3000)
Delete From faVoucherSub Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2011 and intTransactionTypeID<>3000)
Delete From faVoucherAddress Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2011 and intTransactionTypeID<>3000)
Delete From faTransactionChild Where intTransactionId in (Select intTransactionID From faTransactions Where intVoucherID in 
(Select intVoucherID  From faVouchers Where intFinancialYearID= 2011 and intTransactionTypeID<>3000))


Delete From faTransactions Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2011 and intTransactionTypeID<>3000)

Delete From faIDemandAddress Where numDemandID in ( Select numDemandID From faIDemandTBL 
Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2011 and intTransactionTypeID<>3000))

Delete From faIDemandChild Where numDemandID in ( Select numDemandID From faIDemandTBL 
Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2011 and intTransactionTypeID<>3000))

Delete From faIDemandTBL Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2011 and intTransactionTypeID<>3000)

Delete  From faVouchers Where intFinancialYearID= 2011 and intTransactionTypeID<>3000

--2012


Delete From faVoucherChild Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2012 and intTransactionTypeID<>3000)
Delete From faVoucherSub Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2012 and intTransactionTypeID<>3000)
Delete From faVoucherAddress Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2012 and intTransactionTypeID<>3000)
Delete From faTransactionChild Where intTransactionId in (Select intTransactionID From faTransactions Where intVoucherID in 
(Select intVoucherID  From faVouchers Where intFinancialYearID= 2012 and intTransactionTypeID<>3000))


Delete From faTransactions Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2012 and intTransactionTypeID<>3000)

Delete From faIDemandAddress Where numDemandID in ( Select numDemandID From faIDemandTBL 
Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2012 and intTransactionTypeID<>3000))

Delete From faIDemandChild Where numDemandID in ( Select numDemandID From faIDemandTBL 
Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2012 and intTransactionTypeID<>3000))

Delete From faIDemandTBL Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2012 and intTransactionTypeID<>3000)

Delete  From faVouchers Where intFinancialYearID= 2012 and intTransactionTypeID<>3000

--2013
Delete From faVoucherChild_Mirror 
Delete From faVouchers_Mirror 
Delete From faVoucherAddress_Mirror 
Delete From faVoucherSub_Mirror 

alter table faVoucherChild disable trigger all
alter table faVouchers disable trigger all
alter table faVoucherAddress disable trigger all
alter table faVoucherSub disable trigger all
alter table faTransactionChild disable trigger all
alter table faTransactions disable trigger all

Delete From faVoucherChild Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2013 and intTransactionTypeID<>3000)
Delete From faVoucherSub Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2013 and intTransactionTypeID<>3000)
Delete From faVoucherAddress Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2013 and intTransactionTypeID<>3000)

Delete From faTransactionChild Where intTransactionId in (Select intTransactionID From faTransactions Where intVoucherID in 
(Select intVoucherID  From faVouchers Where intFinancialYearID= 2013 and intTransactionTypeID<>3000))

Delete From faTransactions Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2013 and intTransactionTypeID<>3000)

Delete From faIDemandAddress Where numDemandID in ( Select numDemandID From faIDemandTBL 
Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2013 and intTransactionTypeID<>3000))

Delete From faIDemandChild Where numDemandID in ( Select numDemandID From faIDemandTBL 
Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2013 and intTransactionTypeID<>3000))

Delete From faIDemandTBL Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2013 and intTransactionTypeID<>3000)

Delete  From faVouchers Where intFinancialYearID= 2013 and intTransactionTypeID<>3000

---2014
Delete From faVoucherChild Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2014 and intTransactionTypeID<>3000)
Delete From faVoucherSub Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2014 and intTransactionTypeID<>3000)
Delete From faVoucherAddress Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2014 and intTransactionTypeID<>3000)

Delete From faTransactionChild Where intTransactionId in (Select intTransactionID From faTransactions Where intVoucherID in 
(Select intVoucherID  From faVouchers Where intFinancialYearID= 2014 and intTransactionTypeID<>3000))

Delete From faTransactions Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2014 and intTransactionTypeID<>3000)

Delete From faIDemandAddress Where numDemandID in ( Select numDemandID From faIDemandTBL 
Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2014 and intTransactionTypeID<>3000))

Delete From faIDemandChild Where numDemandID in ( Select numDemandID From faIDemandTBL 
Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2014 and intTransactionTypeID<>3000))

Delete From faIDemandTBL Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2014 and intTransactionTypeID<>3000)

Delete  From faVouchers Where intFinancialYearID=2014 and intTransactionTypeID<>3000



---2015
Delete From faVoucherChild Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2015 and intTransactionTypeID<>3000)
Delete From faVoucherSub Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2015 and intTransactionTypeID<>3000)
Delete From faVoucherAddress Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2015 and intTransactionTypeID<>3000)

Delete From faTransactionChild Where intTransactionId in (Select intTransactionID From faTransactions Where intVoucherID in 
(Select intVoucherID  From faVouchers Where intFinancialYearID= 2015 and intTransactionTypeID<>3000))

Delete From faTransactions Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2015 and intTransactionTypeID<>3000)

Delete From faIDemandAddress Where numDemandID in ( Select numDemandID From faIDemandTBL 
Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2015 and intTransactionTypeID<>3000))

Delete From faIDemandChild Where numDemandID in ( Select numDemandID From faIDemandTBL 
Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2015 and intTransactionTypeID<>3000))

Delete From faIDemandTBL Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2015 and intTransactionTypeID<>3000)

Delete  From faVouchers Where intFinancialYearID=2015 and intTransactionTypeID<>3000

---2016
Delete From faVoucherChild Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2016 and intTransactionTypeID<>3000)
Delete From faVoucherSub Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2016 and intTransactionTypeID<>3000)
Delete From faVoucherAddress Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2016 and intTransactionTypeID<>3000)

Delete From faTransactionChild Where intTransactionId in (Select intTransactionID From faTransactions Where intVoucherID in 
(Select intVoucherID  From faVouchers Where intFinancialYearID= 2016 and intTransactionTypeID<>3000))

Delete From faTransactions Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2016 and intTransactionTypeID<>3000)

Delete From faIDemandAddress Where numDemandID in ( Select numDemandID From faIDemandTBL 
Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2016 and intTransactionTypeID<>3000))

Delete From faIDemandChild Where numDemandID in ( Select numDemandID From faIDemandTBL 
Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2016 and intTransactionTypeID<>3000))

Delete From faIDemandTBL Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2016 and intTransactionTypeID<>3000)

Delete  From faVouchers Where intFinancialYearID=2016 and intTransactionTypeID<>3000


---2017
--3428323

Delete From faVoucherChild Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2017  --and intVoucherID <3432898
 and intTransactionTypeID<>3000)

Delete From faVoucherSub Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2017 --and intVoucherID <3491003 
and intTransactionTypeID<>3000)

Delete From faVoucherAddress Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2017 --and intVoucherID <3430000 
and intTransactionTypeID<>3000)

Delete From faTransactionChild Where intTransactionId in (Select intTransactionID From faTransactions Where intVoucherID in 
(Select intVoucherID  From faVouchers Where intFinancialYearID= 2017 and intTransactionTypeID<>3000 and intVoucherID <3251000 ))

Delete From faTransactions Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2017 and intTransactionTypeID<>3000)

Delete From faIDemandAddress Where numDemandID in ( Select numDemandID From faIDemandTBL 
Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2017 and intTransactionTypeID<>3000))
Delete From faIDemandChild Where numDemandID in ( Select numDemandID From faIDemandTBL 
Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2017 and intTransactionTypeID<>3000))

Delete From faIDemandTBL Where intVoucherID in (Select intVoucherID  From faVouchers Where intFinancialYearID= 2017 and intTransactionTypeID<>3000)

Delete  From faVouchers Where intFinancialYearID=2017 and intTransactionTypeID<>3000


Update faVouchers set tnyStatus=4,tnyCancelFlag=1 Where intFinancialYearID=2015
Update faTransactions set tnyStatus=4 Where intFinancialYearID=2015

Update faVouchers set tnyStatus=4,tnyCancelFlag=1 Where intFinancialYearID=2016
Update faTransactions set tnyStatus=4 Where intFinancialYearID=2016

Update faVouchers set tnyStatus=4,tnyCancelFlag=1 Where intFinancialYearID=2017
Update faTransactions set tnyStatus=4 Where intFinancialYearID=2017