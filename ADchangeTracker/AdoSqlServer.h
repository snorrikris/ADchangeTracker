#pragma once

class CAdoSqlServer
{
public:
	CAdoSqlServer();
	~CAdoSqlServer(void);

private:	// Make assignment operator and copy constructor private to prevent
			// accidental object copy.
	void operator=(CAdoSqlServer &source) { assert(FALSE); };
	CAdoSqlServer(CAdoSqlServer &source) { assert(FALSE); };

public:
	BOOL InitSqlConnection(LPWSTR szConnectionString);
	void ExitConnection();	// Called from theCore.Exit()
	_ConnectionPtr OpenSqlConnection();

	BOOL IsSqlConnected() { return m_fIsConnected; }

	BOOL IsSqlConnectionLost() { return m_fConnectionLost; }

	// Call this function every X seconds to retry connect to SQL server.
	// Note - does nothing if m_fConnectionLost is not TRUE.
	BOOL RetrySqlConnection();

	int GetSecondsSinceLastRetry();

	BOOL Call_usp_CheckConnection();

	BOOL Call_usp_ADchgEventEx( LPWSTR pwstrXmlData, const long lNumBytes );

	void LogComError( _com_error &e );

	TCHAR m_szConnectionString[1024];

private:
	void LogProviderError();

	_ConnectionPtr	m_pADO_SQLconnection;
	BOOL			m_fIsInitialized;		// TRUE when m_pADO_SQLconnection initialized.
	BOOL			m_fIsConnected;
	BOOL			m_fConnectionLost;
	int				m_nRetryConnectCount;
	BOOL			m_fRetryingToConnect;
	int				m_nSQLconnUseCount;
	__int64			m_nLastRetryConnectTime;	// Set to current time when connection lost
												// and when RetrySqlConnection called.
	void SetRetryConnectTime();
};
