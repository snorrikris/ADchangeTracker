-- Before you run this script do a find and replace on the DOMAIN name (Ctrl-H).
-- Replace 'DOMAIN' with your Active Directoy DOMAIN name (e.g. 'ACME').

USE [master]
GO
CREATE LOGIN [DOMAIN\SG-ADreports] FROM WINDOWS WITH DEFAULT_DATABASE=[AD_DW]
GO
USE [AD_DW]
GO
CREATE USER [DOMAIN\SG-ADreports] FOR LOGIN [DOMAIN\SG-ADreports]
GO
USE [AD_DW]
GO
ALTER USER [DOMAIN\SG-ADreports] WITH DEFAULT_SCHEMA=[dbo]
GO
USE [AD_DW]
GO
ALTER ROLE [db_datareader] ADD MEMBER [DOMAIN\SG-ADreports]
GO
USE [AD_DW]
GO
GRANT EXECUTE ON [dbo].[usp_GetADevents] TO [DOMAIN\SG-ADreports]
GO


USE [master]
GO
CREATE LOGIN [DOMAIN\ad_audit] FROM WINDOWS WITH DEFAULT_DATABASE=[AD_DW]
GO
USE [AD_DW]
GO
CREATE USER [DOMAIN\ad_audit] FOR LOGIN [DOMAIN\ad_audit]
GO
USE [AD_DW]
GO
ALTER USER [DOMAIN\ad_audit] WITH DEFAULT_SCHEMA=[dbo]
GO
GRANT EXECUTE ON [dbo].[usp_ADchgEventEx] TO [DOMAIN\ad_audit]
GO
GRANT EXECUTE ON [dbo].[usp_CheckConnection] TO [DOMAIN\ad_audit]
GO


