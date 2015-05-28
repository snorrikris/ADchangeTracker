#include "StdAfx.h"
//using namespace std;
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
			theLogSys.Add2LogEsyserr(MOD_NAME, "ADO CreateInstance FAILED", "", hr);
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
			//L"Provider='sqloledb';Data Source='SANDMAN';Initial Catalog='AD_DW';Integrated Security=SSPI;");
		try
		{
			m_pADO_SQLconnection->Open(bstrSqlConn, "", "",
					adConnectUnspecified );
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
	return NULL;
}

//void CAdoSqlServer::CheckSqlConnectionHealth( const char *szLogModuleName )
//{
//	if( m_fSqlExcluded || m_fConnectionLost )
//		return;	// No sql used or connection already lost.
//
//	_ConnectionPtr pSqlConn = OpenSqlConnection( szLogModuleName );
//	if( pSqlConn == NULL )
//		return;
//
//	_RecordsetPtr	pSQLrecs("ADODB.Recordset");
//	std::string strDBver, strCmd = "SELECT Variable FROM Settings WHERE (Name = 'DB_VERSION')";
//	try
//	{
//		pSQLrecs->Open( (const char *)strCmd, 
//			_variant_t((IDispatch *)pSqlConn, true),
//			adOpenForwardOnly, adLockReadOnly, adCmdText);
//		if( !(pSQLrecs->BOF && pSQLrecs->EndOfFile) )	
//		{
//			if( pSQLrecs->Fields->GetItem("Variable")->ActualSize > 0 )
//				strDBver = pSQLrecs->Fields->GetItem("Variable")->Value;
//		}
//	}
//	catch( _com_error &e )
//	{
//		LogComError( e, szLogModuleName );
//	}
//	if( pSQLrecs )
//		if( pSQLrecs->State == adStateOpen )
//			pSQLrecs->Close();
//}

BOOL CAdoSqlServer::RetrySqlConnection()
{
	if( !m_fConnectionLost )
		return TRUE;	// connection not lost.

	// Save time of this retry.
	FILETIME ft;
	GetSystemTimeAsFileTime(&ft);
	m_nLastRetryConnectTime = ((__int64)ft.dwHighDateTime << 32) + ft.dwLowDateTime;

	m_fRetryingToConnect = TRUE;
	m_nSQLconnUseCount++;

	BOOL fResult = FALSE;
	_ConnectionPtr pSqlConn = OpenSqlConnection();
	//if( pSqlConn == NULL )
	//	return;
	if (m_fIsConnected)
	{
		// Call usp_CheckConnection to check SQL connection health.
		if (Call_usp_CheckConnection())
		{
			theLogSys.Add2LogI(MOD_NAME, "SQL server re-connected");
			fResult = TRUE;
		}
	}

	m_fRetryingToConnect = FALSE;
	return fResult;
}

int CAdoSqlServer::GetSecondsSinceLastRetry()
{
	FILETIME sCurTime;
	GetSystemTimeAsFileTime(&sCurTime);
	__int64 curtime = ((__int64)sCurTime.dwHighDateTime << 32) + sCurTime.dwLowDateTime;
	__int64 difftime = curtime - m_nLastRetryConnectTime;
	//__int64 retrytout = 60 * 1000 * 1000 * 10;	// 60 sec * 1000 mS * 1000 uS * 10 (10 * 100nS = 1 uS).
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

			theLogSys.Add2LogE( MOD_NAME, "SQL provider error", szTmp, "" );
        }
    }
}

void CAdoSqlServer::LogComError( _com_error &e )
{
    _bstr_t bstrSource(e.Source());
    _bstr_t bstrDescription(e.Description());
    _bstr_t bstrErrorMessage(e.ErrorMessage());
    //_bstr_t bstrErrorInfo(e.ErrorInfo());

	char szDesc[1024];
	sprintf_s(szDesc, sizeof(szDesc), "Code = %08lx - Code meaning = %s - Source = %s", e.Error(), 
		e.ErrorMessage(), (LPCSTR)bstrSource);

	//BOOL fAlarm = ( m_fRetryingToConnect ) ? FALSE : TRUE;

	theLogSys.Add2LogE( MOD_NAME, "ADO com error", szDesc, (LPCSTR)bstrDescription );

	//vector<string> ProviderErrors = GetProviderErrors(TRUE);
	//int nProvErrors = ProviderErrors.size();
/*
#ifndef DO_NOT_SHOW_MESSAGEBOXES_IN_CAdoSqlServer
	char szMsg[1024];
    // Display Com errors.
	sprintf_s(szMsg, sizeof(szMsg), "Code = %08lx\n"
		"Code meaning = %s\n"
		"Source = %s\n"
		"Description = %s\n", e.Error(), e.ErrorMessage(),
		(LPCSTR)bstrSource, (LPCSTR)bstrDescription);
	string ErrorMsg = szMsg;

	ObjectStateEnum eState = (ObjectStateEnum)m_pADO_SQLconnection->State;
	switch( eState )
	{
	case adStateClosed:
		ErrorMsg += "\nConnection state : adStateClosed\n";
		break;
	case adStateOpen:
		ErrorMsg += "\nConnection state : adStateOpen\n";
		break;
	case adStateConnecting:
		ErrorMsg += "\nConnection state : adStateConnecting\n";
		break;
	case adStateExecuting:
		ErrorMsg += "\nConnection state : adStateExecuting\n";
		break;
	case adStateFetching:
		ErrorMsg += "\nConnection state : adStateFetching\n";
		break;
	}

	if( nProvErrors )
		ErrorMsg += "\nProvider error(s):\n";
	for( int i = 0; i < nProvErrors; i++ )
	{
		ErrorMsg += ProviderErrors[i] + "\n";
	}
	//if( !m_fRetryingToConnect )	// No messagebox when retry connect in progress.
	//	MessageBox( NULL, strMsg, strLogEvent, MB_OK );
	//char szTmp[1024];
 //   // Display Com errors.
	//sprintf_s( szTmp, sizeof(szTmp), "Code = %08lx\n"
	//				"Code meaning = %s\n"
	//				"Source = %s\n"
	//				"Description = %s\n", e.Error(), e.ErrorMessage(),
	//				(LPCSTR)bstrSource, (LPCSTR)bstrDescription );
	//MessageBox( NULL, szTmp, strLogEvent, MB_OK );
#endif
*/
	HRESULT error = e.Error();
	//if (error == 0x80004005)	// Connection lost?
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
			theLogSys.Add2LogE( MOD_NAME, "Connection to SQL server lost", 
				"", "" );
		}
	}
}

