if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spRptJournal]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spRptJournal]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



----------------------------------------------------------------------------------------------------------------------------------------------------
	--  D a t a b a s e		: DB_Finance						            --
	--  C r e a t e d   O n 	:14-Feb-2008 by Dhanya R		           		            --
	--  M o d i f i e d    O n								            --
	--  A p p l i c a t i o n 	: S A A N K H YA  - D o u b l e      E n t r y			            --
	
	-- Calls in Report		: RptJournal  						            --
	--------------------------------------------------------------------------------------------------------------------------------------------------- 

CREATE PROCEDURE spRptJournal
@Fromdate smalldatetime,
@ToDate smalldatetime
 AS
 set dateformat dmy
            SELECT faTransactions.intVoucherNo,
		faTransactions.dtTransactionDate,  
		faAccountHeads.vchAccountHeadCode, 
		faAccountHeads.vchAccountHead, 
		faTransactionChild.intSerialNo, 
		faTransactionChild.tinDebitOrCreditFlag,
		faFunds.vchFundCode,
		faFunds.vchFund,
		faFields.vchFieldCode,
		faFunctions.vchFunctionCode,
		faFunctionaries.vchFunctionaryCode,
		faLocalBody.vchLocalBody,
		faTransactionChild.fltAmount,
		faTransactions.numSubLedgerID,
		DB_Masters..GM_LocalBodyType.chvTypeDescEnglish,
			
			CASE WHEN faTransactionChild.tinDebitOrCreditFlag=1 THEN
				faTransactionChild.fltAmount
			END AS AmountDr,
			CASE WHEN faTransactionChild.tinDebitOrCreditFlag=0 THEN
				faTransactionChild.fltAmount 
			END AS AmountCr,
			faTransactions.vchNarration
	FROM	faTransactions	 INNER JOIN faTransactionChild ON faTransactions.intTransactionID = faTransactionChild.intTransactionID
				 INNER JOIN faFunds ON faFunds.intFundID=faTransactions.intFundID	 
				 LEFT OUTER JOIN  faFields ON faFields.intFieldID=faTransactions.intFieldID 	
				 INNER JOIN  faFunctions ON faFunctions.intFunctionID=faTransactions.intFunctionID 
				 INNER JOIN  faFunctionaries ON faFunctionaries.intFunctionaryID=faTransactions.intFunctionaryID  
				 INNER JOIN  faAccountHeads ON faTransactionChild.intAccountHeadID = faAccountHeads.intAccountHeadID
				 INNER JOIN  faLocalBody ON faTransactions.intLocalBodyID=faLocalBody.intLocalBodyID
				 INNER JOIN   DB_Masters..GM_LocalBodyType ON 
			                    	           DB_Masters..GM_LocalBodyType.tnyLBTypeID=faLocalBody.intCategoryID
           WHERE  (faTransactions.intGroupID = 40 )  and  
	        faTransactions.dtTransactionDate  >@Fromdate  and 
	        faTransactions.dtTransactionDate < @Todate
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

