USE [AD_DW]
GO

ALTER TABLE [dbo].[ADevents] DROP CONSTRAINT [PK_ADevents];

ALTER TABLE [dbo].[ADevents] ALTER COLUMN [EventRecordID] bigint not null;

ALTER TABLE [dbo].[ADevents] ADD CONSTRAINT [PK_ADevents] PRIMARY KEY NONCLUSTERED 
(
	[SourceDC] ASC,
	[EventRecordID] ASC
) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF,
	ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY];

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- ======================================================================================
-- Author:		Snorri Kristjánsson
-- Create date: 16.05.2015
-- Description:	Receives AD changes event data as XML.
-- Called from ADchangeTracker service that runs on domain controllers.
-- Data is extracted from the XML and stored in table columns. The XML data is also stored
-- unchanged.
-- Column:			XML data XPath:
-- [SourceDC]		/Event/System/Computer
-- [EventRecordID]	/Event/System/EventRecordID
-- [EventTime]		/Event/System/TimeCreated/@SystemTime
-- [EventID]		/Event/System/EventID
-- [ObjClass]		**1
-- [Target]			**2
-- [Changes]		**3
-- [ModifiedBy]		/Event/EventData/Data[@Name="SubjectDomainName"]
--					+ '\' + /Event/EventData/Data[@Name="SubjectUserName"]
--
-- **1 [ObjClass] is set depending on EventID:
-- [ObjClass] = 'user' when EventID = 4738, 4740, 4720, 4725, 4724, 4723 OR 4722, 4767.
--
-- [ObjClass] = 'group' when EventID = 4728, 4732, 4733 OR 4756.
--
-- [ObjClass] = 'unknown' when EventID = 4781.
--
-- [ObjClass] = Data from XML, XPath: /Event/EventData/Data[@Name="ObjectClass"] 
--              when EventID = 5136, 5137, 5139, 5141.
--
-- **2 [Target] is set depending on EventID:
-- [Target] = Data from XML, XPath: /Event/EventData/Data[@Name="TargetDomainName"]
--                                  + '\' +/Event/EventData/Data[@Name="TargetUserName"]
--            when EventID = 4738, 4740, 4725, 4724, 4723 OR 4722, 4720, 4732, 4733, 4781, 4728, 4756, 4767.
--
-- [Target] = Data from XML, XPath: /Event/EventData/Data[@Name="ObjectDN"]
--            when EventID = 5136, 5137 OR 5141.
--
-- [Target] = Data from XML, XPath: /Event/EventData/Data[@Name="OldObjectDN"]
--            when EventID = 5139.
--
-- [Target] = Data from XML, XPath: /Event/EventData/Data[@Name="SubjectDomainName"]
--                                  + '\' + /Event/EventData/Data[@Name="SubjectDomainName"]
--            when EventID = 4740.
--
-- **3 [Changes] is set depending on EventID:
-- [Changes] = 'NewTargetUserName: ' + Data from XML, XPath: /Event/EventData/Data[@Name="NewTargetUserName"]
--            when EventID = 4781
--
-- [Changes] = 'MemberName: ' + Data from XML, XPath: /Event/EventData/Data[@Name="MemberName"]
--            when EventID = 4728 OR 4756
--
-- [Changes] = 'MemberSID: ' + Data from XML, XPath: /Event/EventData/Data[@Name="MemberSid"]
--            when EventID = 4732, 4733
--
-- [Changes] = '(Value Added) ' OR '(Value Deleted) ' 
--             + Data from XML, XPath: /Event/EventData/Data[@Name="AttributeLDAPDisplayName"]
--             + ': ' + /Event/EventData/Data[@Name="AttributeValue"]
--            when EventID = 5136
--
-- [Changes] = 'NewObjectDN: ' + Data from XML, XPath: /Event/EventData/Data[@Name="NewObjectDN"]
--            when EventID = 5139
--
-- [Changes] = 'Calling computer: ' + Data from XML, XPath: /Event/EventData/Data[@Name="TargetDomainName"]
--            when EventID = 4740
--
-- ======================================================================================
ALTER PROCEDURE [dbo].[usp_ADchgEventEx]
	@XmlData nvarchar(max)
AS
BEGIN
	SET NOCOUNT ON;

	DECLARE @x XML = @XmlData;

	-- Get EventRecordID and SourceDC from XML data.
	DECLARE @EventRecordID bigint;
	WITH XMLNAMESPACES ( DEFAULT 'http://schemas.microsoft.com/win/2004/08/events/event')
		SELECT @EventRecordID = @x.value('(/Event/System/EventRecordID)[1]', 'bigint');
	DECLARE @SourceDC nvarchar(128);
	WITH XMLNAMESPACES ( DEFAULT 'http://schemas.microsoft.com/win/2004/08/events/event')
		SELECT @SourceDC = @x.value('(/Event/System/Computer)[1]', 'nvarchar(128)');

	-- Early exit if event already processed (exists in table).
	IF EXISTS(SELECT EventRecordID FROM dbo.ADevents 
		WHERE EventRecordID = @EventRecordID AND SourceDC = @SourceDC)
		RETURN;

	DECLARE @EventID int;
	WITH XMLNAMESPACES ( DEFAULT 'http://schemas.microsoft.com/win/2004/08/events/event')
		SELECT @EventID = @x.value('(/Event/System/EventID)[1]', 'int'); -- AS EventID

	DECLARE @ObjClass nvarchar(128), @Target nvarchar(256) = '', @Changes nvarchar(256) = '';

	-- Set @ObjClass depending on EventID:
	SELECT @ObjClass = 'user' WHERE @EventID IN (4740, 4738, 4725, 4724, 4723, 4722, 4720, 4767);
	SELECT @ObjClass = 'unknown' WHERE @EventID IN (4781);
	SELECT @ObjClass = 'group' WHERE @EventID IN (4728, 4732, 4733, 4756);
	WITH XMLNAMESPACES ( DEFAULT 'http://schemas.microsoft.com/win/2004/08/events/event')
		SELECT @ObjClass = @x.value('(/Event/EventData/Data[@Name="ObjectClass"])[1]', 'nvarchar(64)')
		WHERE @EventID IN (5136, 5137, 5139, 5141);

	-- Set @Target depending on EventID:
	WITH XMLNAMESPACES ( DEFAULT 'http://schemas.microsoft.com/win/2004/08/events/event')
		SELECT @Target = @x.value('(/Event/EventData/Data[@Name="SubjectDomainName"])[1]', 'nvarchar(64)') + '\' 
			+ @x.value('(/Event/EventData/Data[@Name="TargetUserName"])[1]', 'nvarchar(64)')
		WHERE @EventID IN (4740);
	WITH XMLNAMESPACES ( DEFAULT 'http://schemas.microsoft.com/win/2004/08/events/event')
		SELECT @Target = @x.value('(/Event/EventData/Data[@Name="TargetDomainName"])[1]', 'nvarchar(64)') + '\' 
			+ @x.value('(/Event/EventData/Data[@Name="TargetUserName"])[1]', 'nvarchar(64)')
		WHERE @EventID IN (4738, 4725, 4724, 4723, 4722, 4720, 4728, 4732, 4733, 4756, 4767);
	WITH XMLNAMESPACES ( DEFAULT 'http://schemas.microsoft.com/win/2004/08/events/event')
		SELECT @Target = @x.value('(/Event/EventData/Data[@Name="TargetDomainName"])[1]', 'nvarchar(64)') + '\' 
			+ @x.value('(/Event/EventData/Data[@Name="OldTargetUserName"])[1]', 'nvarchar(64)')
		WHERE @EventID IN (4781);
	WITH XMLNAMESPACES ( DEFAULT 'http://schemas.microsoft.com/win/2004/08/events/event')
		SELECT @Target = @x.value('(/Event/EventData/Data[@Name="ObjectDN"])[1]', 'nvarchar(128)')
		WHERE @EventID IN (5136, 5137, 5141);
	WITH XMLNAMESPACES ( DEFAULT 'http://schemas.microsoft.com/win/2004/08/events/event')
		SELECT @Target = @x.value('(/Event/EventData/Data[@Name="OldObjectDN"])[1]', 'nvarchar(128)')
	WHERE @EventID IN (5139);

	-- Set @Changes depending on EventID:
	WITH XMLNAMESPACES ( DEFAULT 'http://schemas.microsoft.com/win/2004/08/events/event')
		SELECT @Changes = 'Calling computer: ' 
			+ @x.value('(/Event/EventData/Data[@Name="TargetDomainName"])[1]', 'nvarchar(128)')
		WHERE @EventID IN (4740);
	WITH XMLNAMESPACES ( DEFAULT 'http://schemas.microsoft.com/win/2004/08/events/event')
		SELECT @Changes = 'NewTargetUserName: ' 
			+ @x.value('(/Event/EventData/Data[@Name="NewTargetUserName"])[1]', 'nvarchar(128)')
		WHERE @EventID IN (4781);
	WITH XMLNAMESPACES ( DEFAULT 'http://schemas.microsoft.com/win/2004/08/events/event')
		SELECT @Changes = 'MemberName: ' 
			+ @x.value('(/Event/EventData/Data[@Name="MemberName"])[1]', 'nvarchar(128)')
		WHERE @EventID IN (4728, 4756);
	WITH XMLNAMESPACES ( DEFAULT 'http://schemas.microsoft.com/win/2004/08/events/event')
		SELECT @Changes = 'MemberSID: ' 
			+ @x.value('(/Event/EventData/Data[@Name="MemberSid"])[1]', 'nvarchar(128)')
		WHERE @EventID IN (4732, 4733);
	WITH XMLNAMESPACES ( DEFAULT 'http://schemas.microsoft.com/win/2004/08/events/event')
		SELECT @Changes = 'NewObjectDN: '
			+ @x.value('(/Event/EventData/Data[@Name="NewObjectDN"])[1]', 'nvarchar(128)')
		WHERE @EventID IN (5139);
	IF @EventID = 5136
	BEGIN
		DECLARE @OpType nvarchar(32);
		WITH XMLNAMESPACES ( DEFAULT 'http://schemas.microsoft.com/win/2004/08/events/event')
			SELECT @OpType = @x.value('(/Event/EventData/Data[@Name="OperationType"])[1]', 'nvarchar(32)');
		IF @OpType = '%%14674' 
			SET @OpType = 'Value Added';
		ELSE IF @OpType = '%%14675' 
			SET @OpType = 'Value Deleted';

		WITH XMLNAMESPACES ( DEFAULT 'http://schemas.microsoft.com/win/2004/08/events/event')
			SELECT @Changes = '(' + @OpType + ') ' 
				+ @x.value('(/Event/EventData/Data[@Name="AttributeLDAPDisplayName"])[1]', 'nvarchar(128)') 
				+ ': ' + @x.value('(/Event/EventData/Data[@Name="AttributeValue"])[1]', 'nvarchar(128)');
	END

	-- Insert new row into ADevents table.
	;WITH XMLNAMESPACES (
	  default 'http://schemas.microsoft.com/win/2004/08/events/event'
	)
	,[Event] AS
	(
	SELECT @x.value('(/Event/System/TimeCreated/@SystemTime)[1]', 'datetime2') AS EventTime
		,@x.value('(/Event/EventData/Data[@Name="SubjectDomainName"])[1]', 'nvarchar(64)') + '\' 
			+ @x.value('(/Event/EventData/Data[@Name="SubjectUserName"])[1]', 'nvarchar(64)') AS ModifiedBy
		,@x AS EventXml
	)
	INSERT INTO dbo.ADevents
		SELECT @SourceDC, @EventRecordID, e.EventTime, @EventID AS EventID, 
		@ObjClass AS ObjClass, @Target AS [Target], @Changes AS [Changes], e.ModifiedBy, 
		e.EventXml 
	FROM [Event] e
END
GO
