Declare   @dtFromDate as smalldatetime,	@dtToDate as 	smalldatetime
Set @dtFromDate='1/Mar/2017'
set @dtToDate='1/mar/2017'

SELECT
	SUM(fltAmount)AS Amount,CONVERT(SMALLDATETIME,CONVERT(varchar(11),dtDate)) AS dtdate,numLocationID
	FROM 
	faVouchers
	LEFT JOIN faSeats ON faVouchers.numSeatID = faSeats.numSeatID
	WHERE tnyCancelflag=0 AND tnyVoucherTypeID=10 
	AND dtDate BETWEEN CONVERT(SMALLDATETIME,CONVERT(char(11),@dtFromDate)) AND CONVERT(SMALLDATETIME,CONVERT(char(11),@dtToDate))
	--AND faSeats.intGroupID in (9,10)
	AND isNull(intTransactionTypeID,0) NOT IN(134,135,136,137,75,112,1141,1151,1161,1171,1181,1191,1201)
	AND faVouchers.intInstrumentTypeID=1
	GROUP BY CONVERT(SMALLDATETIME,CONVERT(varchar(11),dtDate)),numLocationID
	ORDER BY CONVERT(SMALLDATETIME,CONVERT(varchar(11),dtDate))


Select sum(faVouchers.fltAmount),dtdate,numLocationID From faVouchers
	             Inner Join faCounters On faCounters.intCounterID = faVouchers.intCounterID
		     inner Join faTransactionType on faTransactionType.intTransactionTypeID=faVouchers.intTransactionTypeID
	             Where faCounters.intSectionID=99 And faTransactionType.intGroupID=10 And
	             dtDate BetWeen Convert(smallDateTime,Convert(varchar(11),@dtFromDate)) And Convert(smallDateTime,Convert(varchar(11),@dtToDate))
	             And ISNULL(tnyCancelFlag,0)=0
	And ISNULL(tnyReversed,0)=0 --and numLocationID=4016701

		     group by dtdate,numLocationID Order by dtdate


--Declare   @dtFromDate as smalldatetime,	@dtToDate as 	smalldatetime
Set @dtFromDate='14/Mar/2017'
set @dtToDate='14/mar/2017'
Select faVouchers.fltAmount,dtdate,* From faVouchers
	             Inner Join faCounters On faCounters.intCounterID = faVouchers.intCounterID
		     inner Join faTransactionType on faTransactionType.intTransactionTypeID=faVouchers.intTransactionTypeID
	             Where faCounters.intSectionID=99 And  faTransactionType.intGroupID=10 And
	             dtDate BetWeen Convert(smallDateTime,Convert(varchar(11),@dtFromDate)) And Convert(smallDateTime,Convert(varchar(11),@dtToDate))
	             And ISNULL(tnyCancelFlag,0)=0 and numLocationID=4016701
	
		     --group by dtdate Order by dtdate
Order by intVoucherNo 






