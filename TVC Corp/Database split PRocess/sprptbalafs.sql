--SpGETAFSBL 4,2011
---SpRptBalanceSheet '31/Mar/2011'
--CREATE Procedure SpGETAFSBL 
	Declare @Month  as int ,           
	 @intYearID as int
--AS
set @Month=4
set @intYearID=2011
	Declare @PreMonth as int
	Declare @PreYearID as int
	DECLARE @vchLBTitle varchar (100)
	Declare @FinStartDate smallDateTime
	Declare @fltInEx numeric(18,2)


	Declare @genfund numeric(18,2)
	Declare @LBType int
	SELECT @vchLBTitle=chvtitle,@LBType=tnyLBTypeID FROM faLBSettings
	if @Month=1  
		set @PreMonth=12  else set @PreMonth=@Month-1
	if @Month=4  set @PreYearID=@intYearID-1 else set @PreYearID=@intYearID
    --    Select @fltInEx=dbo.GetAccountBalance(Null,Null,'3109001%','%')+dbo.GetAccountBalance(Null,@FinStartDate-1,'%','1')+dbo.GetAccountBalance(Null,@FinStartDate-1,'%','2')
	--Select @fltInEx=@fltInEx+dbo.GetAccountBalance(@FinStartDate,@dtDate,'%','1')
	--Select @fltInEx=@fltInEx-dbo.GetAccountBalance(@FinStartDate,@dtDate,'%','2')
	--Select @genfund=dbo.GetAccountBalance(Null,Null,'3101%','%')
---select @PreYearID,@PreMonth
	--Select vchscheduletitle,vchmajoraccountheadcode,vchScheduleGroup,accountheadcode,Accounts,sum(transactionamount) From (
      	SELECT          
                     CASE 
                                WHEN faMajorAccountHeads.vchMajorAccountheadCode like '4%' or faMajorAccountHeads.vchMajorAccountheadCode like '3%' Then
                                            faschedulereports.vchscheduletitle
                                ELSE
                                            'B-1'
                     END AS [vchscheduletitle],     
                     CASE 
                                WHEN faMajorAccountHeads.vchMajorAccountheadCode like '4%' or faMajorAccountHeads.vchMajorAccountheadCode like '3%' Then
                                            famajoraccountheads.vchmajoraccountheadcode
                                ELSE
                                            '310000000'
                     END AS [vchmajoraccountheadcode],
                                       CASE
                                WHEN faMajorAccountHeads.vchMajorAccountheadCode like '4%' or faMajorAccountHeads.vchMajorAccountheadCode like '3%' Then 
                                            faSchedulegroups.vchScheduleGroup
                          ELSE
                                'Reserve& Surplus'
                     END AS [vchScheduleGroup],                    
                     CASE 
                                WHEN faMajorAccountHeads.vchMajorAccountheadCode like '4%'Then
                                            'ASSETS'
                                ELSE
                                            'LIABILITIES'
                     END AS [accountheadtype],
                     CASE 
                                WHEN faMajorAccountHeads.vchMajorAccountheadCode like '4%' or faMajorAccountHeads.vchMajorAccountheadCode like '3%' Then
                                          faMajorAccountHeads.vchMajorAccounthead
                                ELSE
                                            'Muncipal (General) Fund [Code No 310]'
                     END AS [Accounts],
			Sum(CASE  
			WHEN faAccountHeads.TinType in(4) THEN
				faTransactionChild.fltAmount*((faTransactionChild.TinDebitOrCreditFlag*2)-1)
			WHEN faAccountHeads.TinType in(3) THEN
				faTransactionChild.fltAmount*((faTransactionChild.TinDebitOrCreditFlag*-2)+1)
			ELSE
				faTransactionChild.fltAmount*((tinDebitOrCreditFlag*-2)+1)
		END) AS [transactionamount],
		Case When @LBType=3 Or @LBType=4 Then
			CASE 
				WHEN( faAccountHeads.vchAccountheadCode like '4%' or faAccountHeads.vchAccountheadCode like '3%' )  AND  faAccountHeads.vchAccountheadCode<>  '310900100'  THEN
					faaccountheads.vchaccountheadcode
				WHEN (faAccountHeads.vchAccountheadCode like '4%' or faAccountHeads.vchAccountheadCode like '3%' )  AND  faAccountHeads.vchAccountheadCode=  '310900100' THEN
					'310100100'
				ELSE
					'310900100'
			END
		ELSE
			CASE 
				WHEN( faAccountHeads.vchAccountheadCode like '4%' or faAccountHeads.vchAccountheadCode like '3%' )  AND  faAccountHeads.vchAccountheadCode<>  '310900101'  THEN
					faaccountheads.vchaccountheadcode
				WHEN (faAccountHeads.vchAccountheadCode like '4%' or faAccountHeads.vchAccountheadCode like '3%' )  AND  faAccountHeads.vchAccountheadCode=  '310900101' THEN
					'310100101'
				ELSE
					'310900101'
			END
		END AS [accountheadcode],
	      /* Case When @LBType=3 Or @LBType=4 Then	
			CASE 
				WHEN (faAccountHeads.vchAccountheadCode like '4%' or faAccountHeads.vchAccountheadCode like '3%' )  AND  faAccountHeads.vchAccountheadCode<>  '310900100'  THEN
					faAccountHeads.vchAccounthead
				WHEN (faAccountHeads.vchAccountheadCode like '4%' or faAccountHeads.vchAccountheadCode like '3%' )  AND  faAccountHeads.vchAccountheadCode=  '310900100' THEN
					'General Fund'
				ELSE
					'Excess of Income Over Expenditure'
			END
		ELSE
			CASE 
				WHEN (faAccountHeads.vchAccountheadCode like '4%' or faAccountHeads.vchAccountheadCode like '3%' )  AND  faAccountHeads.vchAccountheadCode<>  '310900101'  THEN
					faAccountHeads.vchAccounthead
				WHEN (faAccountHeads.vchAccountheadCode like '4%' or faAccountHeads.vchAccountheadCode like '3%' )  AND  faAccountHeads.vchAccountheadCode=  '310900101' THEN
					'General Fund'
				ELSE
					'Excess of Income Over Expenditure'
			END
		END AS [AccountHead],*/
		CASE 
			WHEN faAccountHeads.vchAccountheadCode like '4%' or faAccountHeads.vchAccountheadCode like '3%' THEN
				faschedulereports.intScheduleReportID
			ELSE
				20
		END As [ReportID]
                           /* SUM(CASE 
                                WHEN faAccountHeads.TinType in(4) THEN
                                            faTransactionChild.fltAmount*((faTransactionChild.TinDebitOrCreditFlag*2)-1)
                                 ELSE
                                            faTransactionChild.fltAmount*((faTransactionChild.TinDebitOrCreditFlag*-2)+1)
                                 END)  AS [transactionamount]  */           

	FROM  faTransactionChild INNER JOIN 
		faTransactions	ON faTransactionChild.intTransactionId=faTransactions.intTransactionId 	INNER JOIN 
		faAccountHeads  ON faTransactionChild.intAccountHeadId=faAccountHeads.intAccountHeadId	INNER JOIN 
		faminoraccountheads ON faaccountheads.intminoraccountheadid=faminoraccountheads.intminoraccountheadid	INNER JOIN 
		famajoraccountheads ON faaccountheads.intmajoraccountheadid=famajoraccountheads.intmajoraccountheadid	LEFT OUTER JOIN 
		faschedulereportheads ON faschedulereportheads.vchaccountheadcode=faminoraccountHeads.vchMinoraccountHeadcode
			        OR faschedulereportheads.vchaccountheadcode=famajoraccountHeads.vchMajoraccountHeadcode
	        		OR faschedulereportheads.vchaccountheadcode=faaccountHeads.vchaccountHeadcode	LEFT OUTER JOIN 
		faschedulereports   ON faschedulereports.intschedulereportid=faschedulereportheads.intschedulereportid	LEFT OUTER JOIN 
		faschedulegroups    ON faschedulegroups.intschedulegroupid=faschedulereports.intschedulegroupid
	WHERE           
		faTransactions.intFinancialYearID=@intYearID and Month(fatransactions.dttransactiondate) =@Month  
		AND (faTransactions.tnyStatus <>4 OR faTransactions.tnyStatus IS NULL)
		AND faScheduleReports.vchScheduleTitle<>'I-1(a)'
	
	GROUP BY                            
	        vchscheduletitle,
	        famajoraccountheads.vchmajoraccounthead,
	        famajoraccountheads.vchmajoraccountheadcode,
	        faSchedulegroups.vchScheduleGroup,
 		faAccountHeads.vchAccountHeadCode,faAccountHeads.vchAccountHead,
		faschedulereports.intScheduleReportID
		
	/*Union All 

	Select vchScheduleTitle vchscheduletitle,vchMajorAccountHeadCODE vchmajoraccountheadcode,vchScheduleGroup,
	vchAccountHeadGroup accountheadcode,vchMajorAccounthead Accounts,fltAmount transactionamount
	From faAFSStmtMonthly 	Inner Join faAFSMonthly On faAFSStmtMonthly.intAfsMonthID=faAFSMonthly.intAfsMonthID
	Where intYearId=@PreYearID and intMonthID=@PreMonth
) A

GROUP BY                            
	        A.vchscheduletitle,
	        A.vchmajoraccountheadcode,
	        A.vchScheduleGroup,
		A.accountheadcode,
		A.Accounts*/
      


GO
