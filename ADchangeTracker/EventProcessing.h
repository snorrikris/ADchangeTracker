#pragma once
#include "AdoSqlServer.h"

// structure to store ObjClass of ignored events with EventID in the range 5136 to 5141
typedef struct tagIgnoreEvents
{
	TCHAR szObjectClass[128];
} IGNORE_EVENTS;

typedef struct tagEvtProcConf
{
	int narrAcceptedEvents[128];					// EventIDs of accepted events.
	int nNumElemAcceptedEvts;						// Number of elements in above array.

	IGNORE_EVENTS sarrIgnoreEvts[16];	// Events to ignore (5136...5141).
	int nNumElemIgnoreEvts;							// Number of elements used in above array.

	TCHAR szConnectionString[1024];	// SQL connection string.

	BOOL fIsVerboseLogging;			// TRUE when log level is verbose.
}	
 EVENT_PROCESSING_CONFIG;

class CEventProcessing
{
public:
	CEventProcessing();
	~CEventProcessing();

	BOOL Init();	// Returns FALSE if service can not start.

	void Start();

	void SetStopSignal();

	EVENT_PROCESSING_CONFIG & GetConfigStruct() { return m_config; }

protected:
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
};

