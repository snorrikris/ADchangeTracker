#include "StdAfx.h"
#include "AdoSqlServer.h"
#include "LogSys.h"

// Name used in Log when 'this' module logs an error.
#define MOD_NAME "SQL connection"

CAdoSqlServer::CAdoSqlServer()
{
	m_fIsConnected = m_fIsInitialized = m_fConnectionLost = m_fRetryingToConnect = FALSE;
	m_nRetryConnectCount = 0;
	m_szConnectionString[0] = 0;
	m_nLastRetryConnectTime = 0;
}

CAdoSqlServer::~CAdoSqlServer(void)
{
}

BOOL CAdoSqlServer::InitSqlConnection(LPWSTR szConnectionString)
{
	// Save connection string.
	StringCchCopy(m_szConnectionString, sizeof(m_szConnectionString), szConnectionString);

	HRESULT hr = S_OK;
	try 
	{ 
		m_fIsInitialized = TRUE;
		hr = m_pADO_SQLconnection.CreateInstance(__uuidof(Connection));
		if( FAILED(hr) )
		{
			theLog.SysErr(MOD_NAME, "ADO CreateInstance FAILED", "", hr);
			return FALSE;
		}
	}
	catch(_com_error &e)
	{
		m_fIsInitialized = FALSE;
		LogProviderError();
		LogComError( e );
		return FALSE;
	}
	return TRUE;
}

void CAdoSqlServer::ExitConnection()
{
	if( m_fIsInitialized )
		m_pADO_SQLconnection.Release();
}

_ConnectionPtr CAdoSqlServer::OpenSqlConnection()
{
	if( m_pADO_SQLconnection->State == adStateClosed )
	{
		BSTR bstrSqlConn = ::SysAllocString(m_szConnectionString);
		try
		{
			m_pADO_SQLconnection->Open(bstrSqlConn, "", "", adConnectUnspecified );
		}
		catch(_com_error &e)
		{
			LogProviderError();
			LogComError( e );
		}
		::SysFreeString(bstrSqlConn);
	}
	if( m_pADO_SQLconnection->State == adStateOpen )
	{
		m_fIsConnected = TRUE;
		m_fConnectionLost = FALSE;
		m_nSQLconnUseCount++;
		if( (m_pADO_SQLconnection->Errors->Count) > 0 )
			m_pADO_SQLconnection->Errors->Clear();
		return m_pADO_SQLconnection;
	}
	else
	{
		// Connect to SQL server failed. Set Connection lost flag.
		SetRetryConnectTime();
		m_fIsConnected = FALSE;
		m_fConnectionLost = TRUE;
	}
	return NULL;
}

BOOL CAdoSqlServer::RetrySqlConnection()
{
	if( !m_fConnectionLost )
		return TRUE;	// connection not lost.

	theLog.Info(MOD_NAME, "Retrying SQL server connect");

	// Save the time of this retry.
	SetRetryConnectTime();

	m_fRetryingToConnect = TRUE;
	m_nSQLconnUseCount++;

	BOOL fResult = FALSE;
	_ConnectionPtr pSqlConn = OpenSqlConnection();
	if (m_fIsConnected)
	{
		// Call usp_CheckConnection to check SQL connection health.
		if (Call_usp_CheckConnection())
		{
			theLog.Info(MOD_NAME, "SQL server re-connected");
			fResult = TRUE;
		}
	}

	m_fRetryingToConnect = FALSE;
	return fResult;
}

void CAdoSqlServer::SetRetryConnectTime()
{
	FILETIME ft;
	GetSystemTimeAsFileTime(&ft);
	m_nLastRetryConnectTime = ((__int64)ft.dwHighDateTime << 32) + ft.dwLowDateTime;
}

int CAdoSqlServer::GetSecondsSinceLastRetry()
{
	FILETIME sCurTime;
	GetSystemTimeAsFileTime(&sCurTime);
	__int64 curtime = ((__int64)sCurTime.dwHighDateTime << 32) + sCurTime.dwLowDateTime;
	__int64 difftime = curtime - m_nLastRetryConnectTime;
	__int64 onesec = 1000 * 1000 * 10;	// 1000 mS * 1000 uS * 10 (10 * 100nS = 1 uS).
	int secs = (int)(difftime / onesec);
	if (secs < 0)
		secs = 0;
	return secs;
}

BOOL CAdoSqlServer::Call_usp_ADchgEventEx(LPWSTR pwstrXmlData, const long lNumBytes)
{
	_ConnectionPtr pSQLConn = OpenSqlConnection();
	if( pSQLConn == NULL )
		return FALSE;
	BOOL fRetval = TRUE;
	try
	{
		_CommandPtr CommandPtr = NULL;
		CommandPtr.CreateInstance(__uuidof(Command));
		CommandPtr->ActiveConnection = pSQLConn;
		CommandPtr->CommandType = adCmdStoredProc;
		CommandPtr->CommandText = _bstr_t("usp_ADchgEventEx");
		CommandPtr->NamedParameters = true;
		_ParameterPtr ParamPtr = CommandPtr->CreateParameter(_bstr_t("@XmlData"), adVarWChar,
			adParamInput, lNumBytes, _bstr_t(pwstrXmlData));
		CommandPtr->Parameters->Append(ParamPtr);
		_RecordsetPtr RecordsetPtr = CommandPtr->Execute(NULL, NULL, adCmdStoredProc);
	}
	catch( _com_error &e )
	{
		LogComError( e );
		fRetval = FALSE;
	}
	if( (pSQLConn->Errors->Count) > 0 )
		fRetval = FALSE;
	return fRetval;
}

BOOL CAdoSqlServer::Call_usp_CheckConnection()
{
	_ConnectionPtr pSQLConn = OpenSqlConnection();
	if( pSQLConn == NULL )
		return FALSE;
	BOOL fRetval = TRUE;
	try
	{
		_CommandPtr CommandPtr = NULL;
		CommandPtr.CreateInstance(__uuidof(Command));
		CommandPtr->ActiveConnection = pSQLConn;
		CommandPtr->CommandType = adCmdStoredProc;
		CommandPtr->CommandText = _bstr_t("usp_CheckConnection");
		_RecordsetPtr RecordsetPtr = CommandPtr->Execute(NULL, NULL, adCmdStoredProc);
	}
	catch( _com_error &e )
	{
		LogComError( e );
		fRetval = FALSE;
	}
	if( (pSQLConn->Errors->Count) > 0 )
		fRetval = FALSE;
	return fRetval;
}

void CAdoSqlServer::LogProviderError()
{
char szTmp[1024];
    // Log Provider Errors from Connection object.
    // pErr is a record object in the Connection's Error collection.
    ErrorPtr  pErr = NULL;

    if( (m_pADO_SQLconnection->Errors->Count) > 0)
    {
        long nCount = m_pADO_SQLconnection->Errors->Count;
        // Collection ranges from 0 to nCount -1.
        for(long i = 0; i < nCount; i++)
        {
            pErr = m_pADO_SQLconnection->Errors->GetItem(i);

			sprintf_s( szTmp, sizeof(szTmp), "Error number: %x\t%s\n", pErr->Number, 
				(LPCSTR)pErr->Description );

			theLog.Error( MOD_NAME, "SQL provider error", szTmp, "" );
        }
    }
}

void CAdoSqlServer::LogComError( _com_error &e )
{
    _bstr_t bstrSource(e.Source());
    _bstr_t bstrDescription(e.Description());
    _bstr_t bstrErrorMessage(e.ErrorMessage());

	char szDesc[1024];
	sprintf_s(szDesc, sizeof(szDesc), "Code = %08lx - Code meaning = %s - Source = %s", e.Error(), 
		e.ErrorMessage(), (LPCSTR)bstrSource);

	theLog.Error( MOD_NAME, "ADO com error", szDesc, (LPCSTR)bstrDescription );

	HRESULT error = e.Error();
	if (error == E_FAIL)	// Connection lost?
	{
		if( m_pADO_SQLconnection->State == adStateOpen )
		{
			m_pADO_SQLconnection->Close();
			m_fIsConnected = FALSE;
			m_fConnectionLost = TRUE;

			// Save of when connection lost.
			FILETIME ft;
			GetSystemTimeAsFileTime(&ft);
			m_nLastRetryConnectTime = ((__int64)ft.dwHighDateTime << 32) + ft.dwLowDateTime;

			m_nRetryConnectCount = 0;
			theLog.Error( MOD_NAME, "Connection to SQL server lost", 
				"", "" );
		}
	}
}

