#include "stdafx.h"
#include "EventProcessing.h"
#include "LogSys.h"

using namespace pugi;

// Name used in Log when 'this' module logs an error.
#define MOD_NAME "Event processing"

//   Entry point for the service - SCM will call this function.
VOID WINAPI SvcMain(DWORD dwArgc, LPTSTR *lpszArgv)
{
	theService.ServiceMain();
}

// Called by SCM whenever a control code is sent to the service using the ControlService function.
// Parameters: dwCtrl - control code
VOID WINAPI SvcCtrlHandler(DWORD dwCtrl)
{
	theService.ServiceCtrlHandler(dwCtrl);
}

/////////////////////////////////////////////////////////////////////////////////////
// The one and only CEventProcessing object - the service code.
CEventProcessing theService;
/////////////////////////////////////////////////////////////////////////////////////

CEventProcessing::CEventProcessing()
{
	m_hSubscription = m_hBookmark = NULL;
	memset(&m_config, 0, sizeof(m_config));
	m_config.fIsVerboseLogging = TRUE;
	m_hEvent_SqlConnLost = m_hEvent_ServiceStop = NULL;

	m_hSvcStatusHandle = 0;
	memset(&m_sSvcStatus, 0, sizeof(m_sSvcStatus));
	m_sSvcStatus.dwServiceType = SERVICE_WIN32_OWN_PROCESS;
	m_dwCheckPoint = 1;
}

CEventProcessing::~CEventProcessing()
{
	assert(m_hSubscription == NULL
		&& m_hBookmark == NULL
		&& m_hEvent_SqlConnLost == NULL
		&& m_hEvent_ServiceStop == NULL);
}

void CEventProcessing::ServiceMain()
{
	theLogSys.Add2LogI(MOD_NAME, "ServiceMain called");

	// Register the handler function for the service
	m_hSvcStatusHandle = RegisterServiceCtrlHandler(SVCNAME, SvcCtrlHandler);
	if (!m_hSvcStatusHandle)
	{
		theLogSys.Add2LogEsyserr(MOD_NAME, "RegisterServiceCtrlHandler failed",
			"", GetLastError());
		return;
	}

	// Report initial status to the SCM
	ReportServiceStatus(SERVICE_START_PENDING, NO_ERROR, 3000);
	theLogSys.Add2LogI(MOD_NAME, "Service starting");

	// Initialize things needed for service start.
	BOOL fIsInitialized = TRUE;

	// Initialize ADO (COM library).
	CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

	// Verify that minimum required settings are present.
	if (_tcslen(m_config.szConnectionString) == 0)
	{
		theLogSys.Add2LogE(MOD_NAME, "SQL connection string missing");
		fIsInitialized = FALSE;
	}
	if (m_config.nNumElemAcceptedEvts == 0)
	{
		theLogSys.Add2LogW(MOD_NAME, "List of accepted EventIDs missing", "This service will do nothing");
		fIsInitialized = FALSE;
	}

	// Create a event object - that will signal when service should stop.
	if ((m_hEvent_ServiceStop = CreateEvent(NULL, TRUE, FALSE, NULL)) == NULL)
	{
		theLogSys.Add2LogEsyserr(MOD_NAME, "Create service stop event failed", "", GetLastError());
		fIsInitialized = FALSE;
	}

	// Create a event object - that will signal if SQL connection is lost.
	if ((m_hEvent_SqlConnLost = CreateEvent(NULL, TRUE, FALSE, NULL)) == NULL)
	{
		theLogSys.Add2LogEsyserr(MOD_NAME, "Create SQL connection lost event failed", "", GetLastError());
		fIsInitialized = FALSE;
	}

	// Initialize SQL server connection (not connecting at this time).
	if (!m_sqlServer.InitSqlConnection(m_config.szConnectionString))
	{
		theLogSys.Add2LogE(MOD_NAME, "Initialize SQL ADO failed");
		fIsInitialized = FALSE;
	}

	if (FALSE == fIsInitialized)	// Initialize failed - service can't start.
	{
		theLogSys.Add2LogE(MOD_NAME, "Initialize service failed", "Service can't start");
		ReportServiceStatus(SERVICE_STOPPED, NO_ERROR, 0);
		return;
	}

	// Report running status when initialization is complete.
	ReportServiceStatus(SERVICE_RUNNING, NO_ERROR, 0);
	theLogSys.Add2LogI(MOD_NAME, "Service running");

	///////////////////////////////////////////////////////////////////////////
	// Call service start function.
	Start();
	// The Start function returns when the service is stopping.

	theLogSys.Add2LogI(MOD_NAME, "Service is stopping");

	StopEventSubscription();

	m_sqlServer.ExitConnection();

	CloseHandle(m_hEvent_SqlConnLost);
	m_hEvent_SqlConnLost = NULL;

	CloseHandle(m_hEvent_ServiceStop);
	m_hEvent_ServiceStop = NULL;

	CoUninitialize();

	ReportServiceStatus(SERVICE_STOPPED, NO_ERROR, 0);
	theLogSys.Add2LogI(MOD_NAME, "Service stopped");
}

void CEventProcessing::ServiceCtrlHandler(DWORD dwCtrl)
{
	theLogSys.Add2LogI(MOD_NAME, "SvcCtrlHandler called");

	// Handle the requested control code. 
	switch (dwCtrl)
	{
	case SERVICE_CONTROL_STOP:
		theLogSys.Add2LogI(MOD_NAME, "SERVICE_CONTROL_STOP command received");
		ReportServiceStatus(SERVICE_STOP_PENDING, NO_ERROR, 1000);

		// Signal the service to stop.
		SetEvent(m_hEvent_ServiceStop);
		//ReportServiceStatus(m_sSvcStatus.dwCurrentState, NO_ERROR, 0);
		return;

	case SERVICE_CONTROL_INTERROGATE:
		break;

	default:
		break;
	}
}

void CEventProcessing::Start()
{
	// Start SQL server connection.
	if (m_sqlServer.OpenSqlConnection())
	{
		// Start Security event log subscription - if connected to SQL.
		StartEventSubscription();
	}
	else
	{
		// Set SQL connection lost event - note: retry connect every 60 seconds.
		SetEvent(m_hEvent_SqlConnLost);
	}

	HANDLE harrEvents[2] = { m_hEvent_ServiceStop, m_hEvent_SqlConnLost };
	DWORD dwNumEvents = sizeof(harrEvents) / sizeof(HANDLE);
	while (TRUE)
	{
		DWORD dwWaitResult = WaitForMultipleObjects(dwNumEvents, harrEvents, FALSE, INFINITE);

		// Check whether to stop the service.
		//if (WaitForSingleObject(m_hEvent_ServiceStop, 0) == WAIT_OBJECT_0)
		if (dwWaitResult == WAIT_OBJECT_0)
		{
			// Stop event is in signaled state.
			theLogSys.Add2LogI(MOD_NAME, "Stop event signaled");
			break;	// Stop the service.
		}

		// m_hEvent_SqlConnLost event is in signaled state until SQL connection regained.
		//if (WaitForSingleObject(m_hEvent_SqlConnLost, 0) == WAIT_OBJECT_0)
		if (dwWaitResult == (WAIT_OBJECT_0 + 1))
		{
			// Stop event subscription - (if needed).
			StopEventSubscription();

			// Retry SQL connection every 60 sec.
			if (m_sqlServer.GetSecondsSinceLastRetry() > 60)
			{
				if (m_sqlServer.RetrySqlConnection())
				{
					// SQL connected again.
					ResetEvent(m_hEvent_SqlConnLost);
					StartEventSubscription();
				}
			}
		}
		Sleep(100);
	}	// endof while loop
}

//   Sets the current service status and reports it to the SCM.
// Parameters:
//   dwCurrentState - The current state (see SERVICE_STATUS)
//   dwWin32ExitCode - The system error code
//   dwWaitHint - Estimated time for pending operation, in milliseconds
void CEventProcessing::ReportServiceStatus(DWORD dwCurrentState,
	DWORD dwWin32ExitCode, DWORD dwWaitHint)
{
	// Fill in the SERVICE_STATUS structure.
	m_sSvcStatus.dwCurrentState = dwCurrentState;
	m_sSvcStatus.dwWin32ExitCode = dwWin32ExitCode;
	m_sSvcStatus.dwWaitHint = dwWaitHint;

	if (dwCurrentState == SERVICE_START_PENDING)
		m_sSvcStatus.dwControlsAccepted = 0;
	else m_sSvcStatus.dwControlsAccepted = SERVICE_ACCEPT_STOP;

	if ((dwCurrentState == SERVICE_RUNNING) || (dwCurrentState == SERVICE_STOPPED))
		m_sSvcStatus.dwCheckPoint = 0;
	else m_sSvcStatus.dwCheckPoint = m_dwCheckPoint++; // inc. on every 'STARTING' update.

	// Report the status of the service to the SCM.
	SetServiceStatus(m_hSvcStatusHandle, &m_sSvcStatus);
}

BOOL CEventProcessing::StartEventSubscription()
{
	theLogSys.Add2LogI(MOD_NAME, "StartEventSubscription called");

	// Get the saved bookmark.
	GetBookmark();

	DWORD status = ERROR_SUCCESS;
	LPWSTR pwsPath = L"Security";
	LPWSTR pwsQuery = L"*";
	BOOL fReturn = TRUE;

	// Subscribe to existing and furture events beginning with the bookmarked event.
	// If the bookmark has not been persisted, pass an empty bookmark and the subscription
	// will begin with the second event that matches the query criteria.
	m_hSubscription = EvtSubscribe(NULL, NULL, pwsPath, pwsQuery, m_hBookmark, (PVOID)this,
		(EVT_SUBSCRIBE_CALLBACK)CEventProcessing::SubscriptionCallback,
		EvtSubscribeStartAfterBookmark);
	if (NULL == m_hSubscription)
	{
		theLogSys.Add2LogEsyserr(MOD_NAME, "EvtSubscribe call failed", "", GetLastError());
		fReturn = FALSE;
	}
	return fReturn;
}

void CEventProcessing::StopEventSubscription()
{
	if (!m_hSubscription)
		return;

	theLogSys.Add2LogI(MOD_NAME, "StopEventSubscription called");

	SaveBookmark();

	if (m_hSubscription)
	{
		EvtClose(m_hSubscription);
		m_hSubscription = NULL;
	}

	if (m_hBookmark)
	{
		EvtClose(m_hBookmark);
		m_hBookmark = NULL;
	}
}

// (static) The callback that receives the events that match the query criteria. 
DWORD WINAPI CEventProcessing::SubscriptionCallback(EVT_SUBSCRIBE_NOTIFY_ACTION action, PVOID pContext,
	EVT_HANDLE hEvent)
{
	assert(pContext != NULL);
	CEventProcessing &thisobj = *(CEventProcessing *)pContext;

	switch (action)
	{
	case EvtSubscribeActionDeliver:
		thisobj.ProcessEvent(hEvent);
		break;

	default:
		theLogSys.Add2LogW(MOD_NAME, "SubscriptionCallback: Unknown action");
	}
	return ERROR_SUCCESS; // The service ignores the returned status.
}

void CEventProcessing::ProcessEvent(EVT_HANDLE hEvent)
{
	DWORD status = ERROR_SUCCESS, dwBufferSize = 0, dwBufferUsed = 0, dwPropertyCount = 0;
	LPWSTR pRenderedContent = NULL;

	if (!EvtRender(NULL, hEvent, EvtRenderEventXml, dwBufferSize, pRenderedContent,
		&dwBufferUsed, &dwPropertyCount))
	{
		if (ERROR_INSUFFICIENT_BUFFER == (status = GetLastError()))
		{
			dwBufferSize = dwBufferUsed;
			pRenderedContent = (LPWSTR)malloc(dwBufferSize);
			if (pRenderedContent)
			{
				EvtRender(NULL, hEvent, EvtRenderEventXml, dwBufferSize,
					pRenderedContent, &dwBufferUsed, &dwPropertyCount);
			}
			else
			{
				theLogSys.Add2LogE(MOD_NAME, "malloc failed in function ProcessEvent");
				goto cleanup;
			}
		}

		if (ERROR_SUCCESS != (status = GetLastError()))
		{
			theLogSys.Add2LogEsyserr(MOD_NAME, "EvtRender failed in function ProcessEvent",
				"", status);
			goto cleanup;
		}
	}

	if (FilterAndSendEventToSql(pRenderedContent, dwBufferUsed + 1))
	{
		// Update bookmark following successful processing of an event.
		if (!EvtUpdateBookmark(m_hBookmark, hEvent))
		{
			status = GetLastError();
			theLogSys.Add2LogEsyserr(MOD_NAME,
				"EvtUpdateBookmark failed in SubscriptionCallback function", "", status);
			goto cleanup;
		}
	}

	if (m_sqlServer.IsSqlConnectionLost())
	{
		SetEvent(m_hEvent_SqlConnLost);
	}

cleanup:
	if (pRenderedContent)
		free(pRenderedContent);
}

BOOL CEventProcessing::FilterAndSendEventToSql(LPWSTR pXML, DWORD dwXMLlen)
{
	BOOL fReturn = TRUE;
	xml_document doc;
	// load document from immutable memory block.
	xml_parse_result result = doc.load_buffer(pXML, dwXMLlen);

	// Get EventRecordID, EventID and ObjectClass (if exists) from Event XML.
	char szEventRecordID[128] = { 0 }, szEventID[128] = { 0 };
	int nEventID = 0;
	xml_node evtrecid = doc.first_element_by_path("/Event/System/EventRecordID");
	if (evtrecid)
	{
		strcpy_s(szEventRecordID, evtrecid.first_child().value());
	}
	xml_node evtid = doc.first_element_by_path("/Event/System/EventID");
	if (evtid)
	{
		strcpy_s(szEventID, evtid.first_child().value());
		nEventID = atoi(szEventID);
	}

	char szOC[128] = { 0 };
	TCHAR szObjClass[128] = { 0 };
	xpath_node objclass = doc.select_node("//Data[@Name='ObjectClass']/text()");
	if (objclass)
	{
		// Make (buffer) safe copy of ObjClass from XML.
		strcpy_s(szOC, objclass.node().value());
		OemToChar(szOC, szObjClass);	// Convert to UNICODE
	}
	if (IsAcceptedEvent(nEventID))
	{
		if (IsIgnoredEvent(nEventID, szObjClass))
		{
			LogInfo("Event ignored", szEventRecordID, szOC);
		}
		else
		{	// Send event to SQL.
			fReturn = m_sqlServer.Call_usp_ADchgEventEx(pXML, dwXMLlen);
			LogInfo("Event sent to SQL", szEventRecordID);
		}
	}
	return fReturn;
}

BOOL CEventProcessing::GetBookmark()
{
	DWORD status = ERROR_SUCCESS;
	LPWSTR pBookmarkXml = NULL;

	// Set pBookmarkXml to the XML string that you persisted in SaveBookmark.
	DWORD dwType;
	WCHAR szBookmark[1024] = { 0 };
	DWORD dataSize = sizeof(szBookmark);
	HKEY hkey = 0;
	long lResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"SOFTWARE\\ADchangeTracker\\",
		0, KEY_READ, &hkey);
	if (ERROR_SUCCESS == lResult)
	{   // read the bookmark value
		lResult = RegQueryValueEx(hkey, L"Bookmark", NULL, &dwType,
			(BYTE*)szBookmark, &dataSize);
		if (ERROR_SUCCESS == lResult)
		{
			pBookmarkXml = szBookmark;
		}
	}
	if (lResult != ERROR_SUCCESS && lResult != ERROR_FILE_NOT_FOUND)
	{
		theLogSys.Add2LogEsyserr(MOD_NAME, "Read Xml event log bookmark from registry failed",
			"", lResult);
	}

	BOOL fReturn = TRUE;
	m_hBookmark = EvtCreateBookmark(pBookmarkXml);
	if (NULL == m_hBookmark)
	{
		theLogSys.Add2LogEsyserr(MOD_NAME, "Create event log bookmark failed",
			"", GetLastError());
		fReturn = FALSE;
	}

	if (hkey)
		RegCloseKey(hkey);

	return fReturn;	// Note - reading registry status is ignored.
}

BOOL CEventProcessing::SaveBookmark()
{
	DWORD status = ERROR_SUCCESS;
	DWORD dwBufferSize = 0;
	DWORD dwBufferUsed = 0;
	DWORD dwPropertyCount = 0;
	LPWSTR pBookmarkXml = NULL;
	HKEY hkey = 0;
	BOOL fReturn = FALSE;

	if (!EvtRender(NULL, m_hBookmark, EvtRenderBookmark, dwBufferSize, pBookmarkXml, 
		&dwBufferUsed, &dwPropertyCount))
	{
		if (ERROR_INSUFFICIENT_BUFFER == (status = GetLastError()))
		{
			dwBufferSize = dwBufferUsed;
			pBookmarkXml = (LPWSTR)malloc(dwBufferSize);
			if (pBookmarkXml)
			{
				EvtRender(NULL, m_hBookmark, EvtRenderBookmark, dwBufferSize, 
					pBookmarkXml, &dwBufferUsed, &dwPropertyCount);
			}
			else
			{
				theLogSys.Add2LogE(MOD_NAME, "malloc failed in SaveBookmark function");
				status = ERROR_OUTOFMEMORY;
				goto cleanup;
			}
		}

		if (ERROR_SUCCESS != (status = GetLastError()))
		{
			theLogSys.Add2LogEsyserr(MOD_NAME, "EvtRender failed in SaveBookmark function",
				"", status);
			goto cleanup;
		}
	}

	// Save bookmark to the registry.
	DWORD dwDisp;
	LONG lResult = RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"SOFTWARE\\ADchangeTracker\\",
		0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hkey, &dwDisp);
	if (ERROR_SUCCESS != lResult)
	{
		theLogSys.Add2LogEsyserr(MOD_NAME, "RegCreateKeyEx failed in SaveBookmark function",
			"", lResult);
		goto cleanup;
	}

	DWORD keyType = 0;
	DWORD dataSize = (wcslen(pBookmarkXml) * 2) + 2;
	lResult = RegSetValueEx(hkey, L"Bookmark", 0, REG_SZ, (BYTE *)pBookmarkXml, dataSize);
	if (lResult != ERROR_SUCCESS)
	{
		theLogSys.Add2LogEsyserr(MOD_NAME, "RegSetValueEx failed in SaveBookmark function",
			"", lResult);
		goto cleanup;
	}
	fReturn = TRUE;

cleanup:
	if (pBookmarkXml)
		free(pBookmarkXml);

	if (hkey)
		RegCloseKey(hkey);

	return fReturn;
}

BOOL CEventProcessing::IsAcceptedEvent(int EventID)
{
	int numElem = sizeof(m_config.narrAcceptedEvents) / sizeof(int);
	for (int i = 0; i < numElem; i++)
	{
		if (EventID == m_config.narrAcceptedEvents[i])
			return TRUE;
	}
	return FALSE;
}

BOOL CEventProcessing::IsIgnoredEvent(int nEventID, TCHAR *szObjectClass)
{
	if (nEventID < 5136 || nEventID > 5141)
		return FALSE;
	int numElem = sizeof(m_config.sarrIgnoreEvts) / sizeof(IGNORE_EVENTS);
	for (int i = 0; i < numElem; i++)
	{
		if (_tcscmp(szObjectClass, m_config.sarrIgnoreEvts[i].szObjectClass) == 0)
			return TRUE;
	}
	return false;
}

void CEventProcessing::LogInfo(const char *szLogEvent,
	const char *szDescription /*= 0*/, const char *szNotes /*= 0*/)
{
	if (!m_config.fIsVerboseLogging)
		return;
	theLogSys.Add2LogI(MOD_NAME, szLogEvent, szDescription, szNotes);
}

