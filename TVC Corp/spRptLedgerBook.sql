if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spRptLedgerBook]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spRptLedgerBook]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




------------------------------------------------------------------------------------- --
	         --   Reports : rptCashBook.rpt                                                --
	         ------------------------------------------------------------------------------------- --
	         --  D a t a b a s e           : DB_Finance                                            --
	         --  C r e a t e d   O n       : 09 June 2008   By   Aiby                      --
	         --  M o d i f i e d    O n    : 17 Nov 2008                       --
	         --  A p p l i c a t i o n     : S A A N K H Y A - D o u b l e      E n t r y          --
	         --                                                                                    --
	         ------------------------------------------------------------------------------------- --
	         --  Calls in Report :- rptCashBook.rpt                                        --
	         --  Calls in Forms  :- frmRptFilterFields
	         --=================================================================================== --
	         --| Number  | Modification Date |   Modified By         |   Line No. From - To           |
	         --|---------|-------------------|-----------------------|--------------------------------|
	         --|         |                   |           |                |
	         --|         |                   |                       |                                |
	         --=======================================================================================
	         
	         
	         CREATE PROCEDURE spRptLedgerBook
	                  @mAccountHeadID    Int         ,
	                  @mStartingDate SmallDateTime   ,
	                  @mEndingDate   SmallDateTime
	         As
	         
	             BEGIN
	                 SELECT @mStartingDate = IsNull ( @mStartingDate , ( SELECT dtStartingDate FROM faFinancialYear WHERE tinCurrentFinancialYearFlag = 1 ) )
	                 SELECT @mEndingDate = IsNull ( @mEndingDate , ( SELECT dtEndingDate FROM faFinancialYear WHERE tinCurrentFinancialYearFlag = 1 ) )
	                 Declare @tinType tinyInt
	                 Declare @mStart smallDateTime
	                 Set @mStart = @mStartingDate
	                 Select @tinType = tinType From faAccountHeads Where intAccountHeadID = @mAccountHeadID
	                 if @tinType < 3 Begin
	                     Set @mStart = '1/Jan/2000'
	                 End
	                 SELECT
	                             A.intTransactionID      ,
	                             A.fltAmount             ,
	                             A.tinDebitOrCreditFlag  ,
	                             A.intAccountHeadID  ,
	                             A.vchAccountHead        ,
	                             A.vchAccountHeadCode    ,
	                             A.intAccountHeadID as AccHeadID ,
	                             A.dtTransactionDate     ,
	                             A.intVoucherID          ,
	                             A.vchGroup          ,
	                             A.intFunctionID         ,
	                             A.intFunctionaryID      ,
	                             A.intFieldID            ,
	                             A.fltOpeningBalance     ,
	                             dbo.faFunctionaries.vchFunctionaryCode  ,
	                             dbo.faFunctions.vchFunctionCode     ,
	                             dbo.faFields.vchFieldCode   ,
	                             SelectionHead           ,
	                             A.tinDrOrCr                     ,         --tinDrOrCr for Fixing error in Ledger Books - Aiby (06-Feb-2008)
	                             A.AcHeadOB          ,
	                             A.vchNarration          ,
	                             A.intBookNo         ,
	                             A.intVoucherNo,
	                             A.intLocalBodyID,
	                             A.vchinstrumentNo,
	                             A.dtInstrumentDate,
	                             A.vchBank,
	                             A.vchBankPlace,
	                             A.vchRefNo,
				     A.vchName,A.vchHouseName,A.vchStreetName,isnull(A.numLocationID,0) numLocationID,A.intTransactionTypeID
	                 From
	                     (
	                     SELECT
	                             dbo.faTransactionChild.intTransactionID     ,
	                             Case when tinDebitOrCreditFlag = 0 then
	                                 dbo.faTransactionChild.fltAmount * -1
	                             Else
	                                 dbo.faTransactionChild.fltAmount
	                             End fltAmount                   ,
	                             dbo.faTransactionChild.tinDebitOrCreditFlag ,
	                             dbo.faTransactionChild.intAccountHeadID ,
	                             dbo.faTransactionChild.fltOpeningBalance,
	                             dbo.faAccountHeads.vchAccountHead       ,
	                             dbo.faAccountHeads.vchAccountHeadCode   ,
	                             dbo.faAccountHeads.intAccountHeadID as AccHeadID,
	                             dbo.faTransactions.dtTransactionDate        ,
	                             dbo.faTransactions.intVoucherID         ,
	                             dbo.faTransactions.vchGroup         ,
	                             dbo.faTransactions.intFunctionID            ,
	                             dbo.faTransactions.intFunctionaryID     ,
	                             dbo.faTransactions.intFieldID           ,
	                             AcHead.vchAccountHeadCode +'   '+AcHead.vchAccountHead as SelectionHead     ,
	                             dbo.faTransactionChild.tinDebitOrCreditFlag as tinDrOrCr,
	                             --AcHead.fltOpeningBalance as AcHeadOB,
	                             0 [AcHeadOB],
	                             dbo.faTransactions.vchNarration,
	                             dbo.faVouchers.intBookNo,
	                             dbo.faVouchers.intVoucherNo,
	                             dbo.faTransactions.intLocalBodyID,
	                             vchinstrumentNo,
	                             dtInstrumentDate,
	                             vchBank,
	                             vchBankPlace,
	                             vchRefNo,
				     vchName,vchHouseName,vchStreetName,numLocationID,faVouchers.intTransactionTypeID
	         
	                     FROM        dbo.faTransactionChild      LEFT JOIN
	                                 dbo.faAccountHeads  ON dbo.faTransactionChild.intByAccountHeadID = dbo.faAccountHeads.intAccountHeadID  RIGHT JOIN
	                             dbo.faTransactions  ON dbo.faTransactions.intTransactionID = dbo.faTransactionChild.intTransactionID    LEFT  JOIN
	                             dbo.faVouchers          ON dbo.faVouchers.intVoucherID = dbo.faTransactions.intVoucherID            INNER JOIN
	                             dbo.faAccountHeads  AcHead  ON  AcHead.intAccountHeadID = @mAccountHeadID
					Left Join faVoucherAddress On faVoucherAddress.intVoucherID=faVouchers.intVoucherID
	                     WHERE       (
	                             dbo.faTransactionChild.intAccountHeadID = @mAccountHeadID   AND
	                             dbo.faTransactionChild.intByAccountHeadID IS Not Null
	                             AND (faTransactions.tnyStatus <> 4 OR faTransactions.tnyStatus IS NULL)
	                             )
	                     Union All
	         
	                         SELECT
	                                 dbo.faTransactionChild.intTransactionID     ,
	                                 Case when tinDebitOrCreditFlag = 0 then
	                                     dbo.faTransactionChild.fltAmount
	                                 Else
	                                     dbo.faTransactionChild.fltAmount * -1
	                                 End fltAmount                   ,
	                                 Case when tinDebitOrCreditFlag = 0 then
	         				1
	                                 Else
	         				0
	                                 End tinDebitOrCreditFlag            ,
	         
	                                 dbo.faTransactionChild.intAccountHeadID ,
	                                 dbo.faTransactionChild.fltOpeningBalance,
	                                 dbo.faAccountHeads.vchAccountHead       ,
	                                 dbo.faAccountHeads.vchAccountHeadCode   ,
	                                 dbo.faAccountHeads.intAccountHeadID as AccHeadID    ,
	                                 dbo.faTransactions.dtTransactionDate        ,
	                                 dbo.faTransactions.intVoucherID         ,
	                                 dbo.faTransactions.vchGroup         ,
	                                 dbo.faTransactions.intFunctionID            ,
	                                 dbo.faTransactions.intFunctionaryID     ,
	                                 dbo.faTransactions.intFieldID           ,
	                                 AcHead.vchAccountHeadCode +'   '+AcHead.vchAccountHead      as SelectionHead        ,
	                                 Case
	                                     When dbo.faTransactionChild.tinDebitOrCreditFlag = 0 Then 1
	                                     Else 0
	                                 End As tinDrOrCr,
	                                 --AcHead.fltOpeningBalance as AcHeadOB,
	                                 0 [AcHeadOB],
	                                 dbo.faTransactions.vchNarration,
	                                 dbo.faVouchers.intBookNo,
	                                 dbo.faVouchers.intVoucherNo,
	                                 dbo.faTransactions.intLocalBodyID,
	                                 vchinstrumentNo,
	                                 dtInstrumentDate,
	                                 vchBank,
	                                 vchBankPlace,
	                                 vchRefNo,
	         			 vchName,vchHouseName,vchStreetName,numLocationID,faVouchers.intTransactionTypeID
	         
	                          FROM           dbo.faTransactionChild  LEFT JOIN
	                                     dbo.faAccountHeads ON dbo.faTransactionChild.intAccountHeadID = dbo.faAccountHeads.intAccountHeadID  RIGHT JOIN
	                                 dbo.faTransactions ON dbo.faTransactions.intTransactionID = dbo.faTransactionChild.intTransactionID  LEFT  JOIN
	                                 dbo.faVouchers          ON dbo.faVouchers.intVoucherID = dbo.faTransactions.intVoucherID         INNER JOIN
	                                 dbo.faAccountHeads AcHead   ON AcHead.intAccountHeadID = @mAccountHeadID
	                          	 Left Join faVoucherAddress On faVoucherAddress.intVoucherID=faVouchers.intVoucherID
				WHERE      (dbo.faTransactionChild.intByAccountHeadID = @mAccountHeadID
	                                 --AND CONVERT ( SmallDateTime,CONVERT( char(11), dbo.faTransactions.dtTransactionDate))  BETWEEN @mStartingDate AND @mEndingDate
	                                 AND (faTransactions.tnyStatus <> 4 OR faTransactions.tnyStatus IS NULL)
	                                 )
	                     )   A   LEFT JOIN
	                     dbo.faFunctions ON dbo.faFunctions.intFunctionID = A.intFunctionID  LEFT JOIN
	                     dbo.faFunctionaries ON dbo.faFunctionaries.intFunctionaryID = A.intFunctionaryID    LEFT JOIN
	                     dbo.faFields ON dbo.faFields.intFieldID = A.intFieldID
	                     WHERE Convert(smallDateTime,Convert(varchar(11),A.dtTransactionDate)) Between Convert(smallDateTime,Convert(varchar(11),@mStartingDate)) AND Convert(smallDateTime,Convert(varchar(11),@mEndingDate)) AND (A.intTransactionID <> 0)
	         
	         
	         
	         Union All
	             (SELECT
	                 0 [intTransactionID],
	                 Sum(fltAmount*((tinDebitOrCreditflag*2)-1)) [fltOpeningBalance],
	                 case when Sum(fltAmount*((tinDebitOrCreditflag*2)-1))>0 then 1 else 0 end,
	                 faTransactionChild.intAccountHeadID,
	                 'Opening Balance' [vchAccountHead]      ,
	                 vchAccountHeadCode  ,
	                 faTransactionChild.intAccountHeadID as AccHeadID    ,
	                 @mStartingDate[dtTransactionDate]       ,
	                 1[intVoucherID]         ,
	                 '' [vchGroup]           ,
	                 0 [intFunctionID]           ,
	                 0 [intFunctionaryID]        ,
	                 0 [intFieldID]          ,
	                 faAccountHeads.fltOpeningBalance        ,
	                 ''[vchFunctionaryCode]  ,
	                 ''[vchFunctionCode]     ,
	                 ''[vchFieldCode]    ,
	                 vchAccountHeadCode +'   '+vchAccountHead [SelectionHead]            ,
	                 Null,--tinDebitOrCredit[tinDrOrCr],
	                 faAccountHeads.fltOpeningBalance[AcHeadOB]          ,
	                 'Opening Balance'[vchNarration]         ,
	                 1[intBookNo]            ,
	                 0[intVoucherNo],
	                 faTransactions.intLocalBodyID,
	                 Null as vchinstrumentNo,
	                 Null as dtInstrumentNo,
	                 Null as vchBank,
	                 Null as vchBankPlace,
	                 Null As vchRefNo,
			 Null As vchName,Null As vchHouseName,Null As vchStreetName,null numLocationID, Null as intTransactionTypeID
	             FROM faTransactions INNER JOIN
	                 faTransactionChild ON   faTransactions.inttransactionID=faTransactionChild.intTransactionID INNER JOIN
	                 faAccountHeads ON   faAccountHeads.intAccountHeadID=faTransactionChild.intAccountHeadID
	         
	             WHERE faTransactionChild.intAccountHeadID = @mAccountHeadID
	                 AND (faTransactions.tnyStatus <> 4 OR faTransactions.tnyStatus IS NULL)
	                 AND dtTransactionDate<convert(smalldatetime,(convert(varchar(11),@mStart)))
	                     OR (faTransactions.intTransactionID=0 AND faTransactionChild.intAccountHeadID = @mAccountHeadID )
	         
	             GROUP BY faAccountHeads.fltOpeningBalance,
	                 faTransactionChild.intAccountHeadID,
	                 faAccountHeads.vchAccountHeadCode,
	                 faAccountHeads.vchAccountHead,
	                 faTransactions.intLocalBodyID)
	         
	         
	             ORDER BY  A.vchGroup,  A.intTransactionID Asc
	         
	         End



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

