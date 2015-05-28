#pragma once
#include "EventProcessing.h"

VOID SvcInstall();
VOID  SvcUninstall();
VOID WINAPI SvcMain(DWORD dwArgc, LPTSTR *lpszArgv);
VOID WINAPI SvcCtrlHandler(DWORD dwCtrl);

class CADchangeTrackerSvc
{
public:
	CADchangeTrackerSvc();
	~CADchangeTrackerSvc() {}

	void ServiceMain();
	void ServiceCtrlHandler(DWORD dwCtrl);

protected:
	SERVICE_STATUS_HANDLE	m_hSvcStatusHandle;
	SERVICE_STATUS			m_sSvcStatus;
	DWORD					m_dwCheckPoint;			// Service start checkpoint.

	void ReportServiceStatus(DWORD dwCurrentState, DWORD dwWin32ExitCode, DWORD dwWaitHint);

	BOOL ReadConfigFile();
	void ProcessConfigFile(BYTE *pFileData, DWORD dwDataLen);
	void ParseConfigFileLine(TCHAR *szLine);
	int ParseAcceptedIDs(TCHAR *szEventIDs, int *pnarrEvents, int nNumElem);
	int ParseIgnoredEvts(TCHAR *szIgnoredEvts, IGNORE_EVENTS *psarrIgnoreEvents, int nNumElem);
	BOOL ParseBoolParam(TCHAR *szParam);

	CEventProcessing m_eventproc;
};

