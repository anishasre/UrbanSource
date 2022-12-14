if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spRptCollectedHead]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spRptCollectedHead]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE Proc spRptCollectedHead	
	(@intAccountHeadID	int,
	 @dtFromDate		smallDateTime,
	 @dtToDate		smallDateTime)
As

     Declare @vchCode as varchar(10)
     Declare @vchHead as varchar(250)
/*Set @intAccountHeadID = 1129
Set @dtFromDate = '1/Mar/2009'
Set @dtToDate = '31/Mar/2009'*/
	Select @vchCode=vchAccountHeadCode,@vchHead=vchAccountHead From faAccountHeads Where intAccountHeadID=@intAccountHeadID
	Select @vchCode headCode,@vchHead HeadName,*
	From faVouchers
	Inner Join faVoucherChild On faVouchers.intVoucherID = faVoucherChild.intVoucherID
	Left Join faVoucherAddress On faVouchers.intVoucherID = faVoucherAddress.intVoucherID
	Inner Join faAccountHeads On faAccountHeads.intAccountHeadID = faVoucherChild.intAccountHeadID
	Inner Join faLBSettings On faLBSettings.intLBID = faVouchers.intLocalBodyID
	Where tnyCancelFlag = 0 
	And DtDate BetWeen Convert(smalldatetime,Convert(varchar(11),@dtFromDate)) And Convert(smalldatetime,Convert(varchar(11),@dtToDate))  and tnyVoucherTypeID = 10 
	And faVouchers.intVoucherID in(	Select intVoucherID From faVoucherChild Where intAccountHeadID = @intAccountHeadID)
	Order by DtDate Asc

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

