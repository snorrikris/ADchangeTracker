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
	theLog.Info(MOD_NAME, "ServiceMain called");

	// Register the handler function for the service
	m_hSvcStatusHandle = RegisterServiceCtrlHandler(SVCNAME, SvcCtrlHandler);
	if (!m_hSvcStatusHandle)
	{
		theLog.SysErr(MOD_NAME, "RegisterServiceCtrlHandler failed", "", GetLastError());
		return;
	}

	// Report initial status to the SCM
	ReportServiceStatus(SERVICE_START_PENDING, NO_ERROR, 3000);
	theLog.Info(MOD_NAME, "Service starting");

	// Initialize things needed for service start.
	BOOL fIsInitialized = TRUE;

	// Initialize ADO (COM library).
	CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

	// Verify that minimum required settings are present.
	if (_tcslen(m_config.szConnectionString) == 0)
	{
		theLog.Error(MOD_NAME, "SQL connection string missing");
		fIsInitialized = FALSE;
	}
	if (m_config.nNumElemAcceptedEvts == 0)
	{
		theLog.Warning(MOD_NAME, "List of accepted EventIDs missing", "This service will do nothing");
		fIsInitialized = FALSE;
	}

	// Create a event object - that will signal when service should stop.
	if ((m_hEvent_ServiceStop = CreateEvent(NULL, TRUE, FALSE, NULL)) == NULL)
	{
		theLog.SysErr(MOD_NAME, "Create service stop event failed", "", GetLastError());
		fIsInitialized = FALSE;
	}

	// Create a event object - that will signal if SQL connection is lost.
	if ((m_hEvent_SqlConnLost = CreateEvent(NULL, TRUE, FALSE, NULL)) == NULL)
	{
		theLog.SysErr(MOD_NAME, "Create SQL connection lost event failed", "", GetLastError());
		fIsInitialized = FALSE;
	}

	// Initialize SQL server connection (not connecting at this time).
	if (!m_sqlServer.InitSqlConnection(m_config.szConnectionString))
	{
		theLog.Error(MOD_NAME, "Initialize SQL ADO failed");
		fIsInitialized = FALSE;
	}

	if (FALSE == fIsInitialized)	// Initialize failed - service can't start.
	{
		theLog.Error(MOD_NAME, "Initialize service failed", "Service can't start");
		ReportServiceStatus(SERVICE_STOPPED, NO_ERROR, 0);
		return;
	}

	// Set days to keep old log files (note setting kicks in next time create new log file is called).
	theLog.SetDaysToKeepOldLogFiles(m_config.nDaysToKeepOldLogFiles);

	// Report running status when initialization is complete.
	ReportServiceStatus(SERVICE_RUNNING, NO_ERROR, 0);
	theLog.Info(MOD_NAME, "Service running");

	///////////////////////////////////////////////////////////////////////////
	// Call service start function.
	Start();
	// The Start function returns when the service is stopping.

	theLog.Info(MOD_NAME, "Service is stopping");

	StopEventSubscription();

	m_sqlServer.ExitConnection();

	CloseHandle(m_hEvent_SqlConnLost);
	m_hEvent_SqlConnLost = NULL;

	CloseHandle(m_hEvent_ServiceStop);
	m_hEvent_ServiceStop = NULL;

	CoUninitialize();

	ReportServiceStatus(SERVICE_STOPPED, NO_ERROR, 0);
	theLog.Info(MOD_NAME, "Service stopped");
}

void CEventProcessing::ServiceCtrlHandler(DWORD dwCtrl)
{
	theLog.Info(MOD_NAME, "SvcCtrlHandler called");

	// Handle the requested control code. 
	switch (dwCtrl)
	{
	case SERVICE_CONTROL_STOP:
		theLog.Info(MOD_NAME, "SERVICE_CONTROL_STOP command received");
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
			theLog.Info(MOD_NAME, "Stop event signaled");
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
	theLog.Info(MOD_NAME, "StartEventSubscription called");

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
		theLog.SysErr(MOD_NAME, "EvtSubscribe call failed", "", GetLastError());
		fReturn = FALSE;
	}
	return fReturn;
}

void CEventProcessing::StopEventSubscription()
{
	if (!m_hSubscription)
		return;

	theLog.Info(MOD_NAME, "StopEventSubscription called");

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
		theLog.Warning(MOD_NAME, "SubscriptionCallback: Unknown action");
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
				theLog.Error(MOD_NAME, "malloc failed in function ProcessEvent");
				goto cleanup;
			}
		}

		if (ERROR_SUCCESS != (status = GetLastError()))
		{
			theLog.SysErr(MOD_NAME, "EvtRender failed in function ProcessEvent", "", status);
			goto cleanup;
		}
	}

	if (FilterAndSendEventToSql(pRenderedContent, dwBufferUsed + 1))
	{
		// Update bookmark following successful processing of an event.
		if (!EvtUpdateBookmark(m_hBookmark, hEvent))
		{
			status = GetLastError();
			theLog.SysErr(MOD_NAME,
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
	LPWSTR pBookmarkXml = NULL;
	BYTE *pFileData = NULL;
	HANDLE hFile = NULL;

	// Read bookmark from a file (in log folder).
	char szBookmarkFile[MAX_PATH];
	strcpy_s(szBookmarkFile, theLog.GetLogPath());
	strcat_s(szBookmarkFile, "Bookmark.bin");
	hFile = CreateFileA(szBookmarkFile, GENERIC_READ, 0, NULL, OPEN_EXISTING,
		FILE_ATTRIBUTE_NORMAL, NULL);
	if (INVALID_HANDLE_VALUE == hFile)
	{
		DWORD dwError = GetLastError();
		if (ERROR_FILE_NOT_FOUND == dwError)
		{
			theLog.Info(MOD_NAME, "Bookmark file not found", szBookmarkFile);
		}
		else
		{
			theLog.SysErr(MOD_NAME, "Read file failed in GetBookmark function",
				szBookmarkFile, dwError);
		}
		hFile = NULL;
		goto no_bookmark_file;
	}
	LARGE_INTEGER nFileSize;
	if (!GetFileSizeEx(hFile, &nFileSize))
	{
		theLog.SysErr(MOD_NAME, "Failed to get size of Bookmark file", szBookmarkFile, GetLastError());
		goto no_bookmark_file;
	}
	// Allocate memory for bookmark + 2 bytes for terminating 0x00
	pFileData = (BYTE *)malloc(nFileSize.LowPart + 2);
	if (NULL == pFileData)
	{
		theLog.Error(MOD_NAME, "Failed to allocate memory for Boommark", szBookmarkFile);
		return FALSE;
		goto no_bookmark_file;
	}
	DWORD dwBytesRead = 0;
	if (!ReadFile(hFile, pFileData, nFileSize.LowPart, &dwBytesRead, NULL))
	{
		theLog.SysErr(MOD_NAME, "Failed to read Bookmark file", szBookmarkFile, GetLastError());
		goto no_bookmark_file;
	}
	*(WORD *)(pFileData + dwBytesRead) = 0;	// Terminate data with 0x00.
	pBookmarkXml = (LPWSTR)pFileData;

no_bookmark_file:	// Note - bookmark is created "blank" if no bookmark file.

	BOOL fReturn = TRUE;
	m_hBookmark = EvtCreateBookmark(pBookmarkXml);
	if (NULL == m_hBookmark)
	{
		theLog.SysErr(MOD_NAME, "Create event log bookmark failed", "", GetLastError());
		fReturn = FALSE;
	}

	if (pFileData)
		free(pFileData);
	if (hFile)
		CloseHandle(hFile);

	return fReturn;
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

	// Render current Bookmark as XML (UNICODE)
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
				theLog.Error(MOD_NAME, "malloc failed in SaveBookmark function");
				status = ERROR_OUTOFMEMORY;
				goto cleanup;
			}
		}

		if (ERROR_SUCCESS != (status = GetLastError()))
		{
			theLog.SysErr(MOD_NAME, "EvtRender failed in SaveBookmark function", "", status);
			goto cleanup;
		}
	}

	// Save bookmark to a file (in log folder).
	char szBookmarkFile[MAX_PATH];
	strcpy_s(szBookmarkFile, theLog.GetLogPath());
	strcat_s(szBookmarkFile, "Bookmark.bin");
	HANDLE hFile = CreateFileA(szBookmarkFile, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS,
		FILE_ATTRIBUTE_NORMAL, NULL);
	if (INVALID_HANDLE_VALUE == hFile)
	{
		theLog.SysErr(MOD_NAME, "Create file failed in SaveBookmark function",
			szBookmarkFile, GetLastError());
		goto cleanup;
	}
	DWORD dwBytesWritten = 0;
	if (!WriteFile(hFile, pBookmarkXml, dwBufferUsed, &dwBytesWritten, NULL))
	{
		theLog.SysErr(MOD_NAME, "Write file failed in SaveBookmark function",
			szBookmarkFile, GetLastError());
	}
	CloseHandle(hFile);

	fReturn = TRUE;	// Save Bookmark OK.

cleanup:
	if (pBookmarkXml)
		free(pBookmarkXml);

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
	theLog.Info(MOD_NAME, szLogEvent, szDescription, szNotes);
}

