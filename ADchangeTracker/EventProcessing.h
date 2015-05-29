#pragma once
#include "AdoSqlServer.h"

// Note - NT service code used is on MSDN: https://msdn.microsoft.com/en-us/library/windows/desktop/bb540475(v=vs.85).aspx

#define SVCNAME			L"ADchangeTracker"
#define SVCDISPNAME		L"Active Directory change tracker"
#define SVCDESCRIPTION	L"Collects selected Active Directory change events into a SQL database."

VOID WINAPI SvcMain(DWORD dwArgc, LPTSTR *lpszArgv);
VOID WINAPI SvcCtrlHandler(DWORD dwCtrl);

// structure to store ObjClass of ignored events with EventID in the range 5136 to 5141
typedef struct tagIgnoreEvents
{
	TCHAR szObjectClass[128];
} IGNORE_EVENTS;

// Configuration settings for service.
typedef struct tagEvtProcConf
{
	int narrAcceptedEvents[128];		// EventIDs of accepted events.
	int nNumElemAcceptedEvts;			// Number of elements in above array.

	IGNORE_EVENTS sarrIgnoreEvts[16];	// Events to ignore (5136...5141).
	int nNumElemIgnoreEvts;				// Number of elements used in above array.

	TCHAR szConnectionString[1024];		// SQL connection string.

	BOOL fIsVerboseLogging;				// TRUE when log level is verbose.
}	
EVENT_PROCESSING_CONFIG;

class CEventProcessing
{
public:
	CEventProcessing();
	~CEventProcessing();

	void ServiceMain();
	void ServiceCtrlHandler(DWORD dwCtrl);

	void Start();

	EVENT_PROCESSING_CONFIG & GetConfigStruct() { return m_config; }

protected:
	void ReportServiceStatus(DWORD dwCurrentState, DWORD dwWin32ExitCode, DWORD dwWaitHint);

	BOOL StartEventSubscription();
	void StopEventSubscription();

	static DWORD WINAPI SubscriptionCallback(EVT_SUBSCRIBE_NOTIFY_ACTION action, 
		PVOID pContext, EVT_HANDLE hEvent);

	void ProcessEvent(EVT_HANDLE hEvent);

	// dwBufferSize = length of XML in bytes - including terminating zero.
	// Returns FALSE if error.
	BOOL FilterAndSendEventToSql(LPWSTR pXML, DWORD dwXMLlen);

	BOOL GetBookmark();
	BOOL SaveBookmark();

	BOOL IsAcceptedEvent(int EventID);
	BOOL IsIgnoredEvent(int nEventID, TCHAR *szObjectClass);

	// Log if log level set to verbose.
	void LogInfo(const char *szLogEvent, const char *szDescription = 0, const char *szNotes = 0);

	EVT_HANDLE m_hSubscription;
	EVT_HANDLE m_hBookmark;
	CAdoSqlServer	m_sqlServer;
	EVENT_PROCESSING_CONFIG m_config;
	HANDLE m_hEvent_SqlConnLost, m_hEvent_ServiceStop;

	// NT service data
	SERVICE_STATUS_HANDLE	m_hSvcStatusHandle;		// Note - the handle does not have to be closed.
	SERVICE_STATUS			m_sSvcStatus;
	DWORD					m_dwCheckPoint;			// Service start checkpoint.
};

////////////////////////////////////////////////////////////////////////////
// Note 'theService' object is instantiated only once and it is global data.
extern CEventProcessing theService;
////////////////////////////////////////////////////////////////////////////
