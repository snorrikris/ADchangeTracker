#include "stdafx.h"
#include "LogSys.h"
#include "ADchangeTracker.h"

// Name used in Log when 'this' module logs an error.
#define MOD_NAME "ADchangeTracker"

int _tmain(int argc, _TCHAR* argv[])
{
	theLog.InitLogSys(GetModuleHandle(NULL));
	theLog.CreateNewLogFile();

	theLog.Info(MOD_NAME, "main called");

	// If command-line parameter is "-install", install the service. 
	// or if command-line parameter is "-uninstall", uninstall the service.
	// Otherwise, the service is probably being started by the SCM.
	if (lstrcmpi(argv[1], L"-install") == 0)
	{
		SvcInstall();
		return 0;
	}
	else if (lstrcmpi(argv[1], L"-uninstall") == 0)
	{
		SvcUninstall();
		return 0;
	}

	// Read service configuration file. Note settings are stored in theService object.
	ReadConfigFile();

	// Connect the main thread of a service process to the service control manager, 
	// which causes the thread to be the service control dispatcher thread for 
	// the calling process.
	SERVICE_TABLE_ENTRY DispatchTable[] = {
		{ SVCNAME, (LPSERVICE_MAIN_FUNCTION)SvcMain },
		{ NULL, NULL } };

	// This call returns when the service has stopped. 
	// The process should simply terminate when the call returns.
	if (!StartServiceCtrlDispatcher(DispatchTable))
	{
		DWORD dwError = GetLastError();
		theLog.SysErr(MOD_NAME, "StartServiceCtrlDispatcher call failed", "", dwError);
		if (ERROR_FAILED_SERVICE_CONTROLLER_CONNECT == dwError)
		{
			theLog.Warning(MOD_NAME, "Attempt to run service as console application");
			printf("This is a service.\n"
				"To install or uninstall the service run a Command Prompt as a Administrator\n"
				"Do: ADchangeTracker -install\n"
				"Or: ADchangeTracker -uninstall\n\n");
		}
	}

	theLog.Info(MOD_NAME, "main ends");
	return 0;
}

//   Installs a service in the SCM database
VOID SvcInstall()
{
	SC_HANDLE hSCManager = 0, hService = 0;
	TCHAR szPath[MAX_PATH];
	if (!GetModuleFileName(GetModuleHandle(NULL), szPath, MAX_PATH))
	{
		theLog.SysErr(MOD_NAME, "Cannot install service", "", GetLastError());
		printf("Cannot install service (%d)\n", GetLastError());
		return;
	}

	hSCManager = OpenSCManager(  // Get a handle to the SCM database. 
		NULL,                    // local computer
		NULL,                    // ServicesActive database 
		SC_MANAGER_ALL_ACCESS);  // full access rights 

	if (NULL == hSCManager)
	{
		theLog.SysErr(MOD_NAME, "OpenSCManager failed", "", GetLastError());
		printf("OpenSCManager failed (%d)\n", GetLastError());
		return;
	}

	hService = CreateService(      // Create the service
		hSCManager,                // SCM database 
		SVCNAME,                   // name of service 
		SVCDISPNAME,               // service name to display 
		SERVICE_ALL_ACCESS,        // desired access 
		SERVICE_WIN32_OWN_PROCESS, // service type 
		SERVICE_DEMAND_START,      // start type 
		SERVICE_ERROR_NORMAL,      // error control type 
		szPath,                    // path to service's binary 
		NULL,                      // no load ordering group 
		NULL,                      // no tag identifier 
		NULL,                      // no dependencies 
		NULL,                      // LocalSystem account 
		NULL);                     // no password 

	if (hService == NULL)
	{
		theLog.SysErr(MOD_NAME, "CreateService failed", "", GetLastError());
		printf("CreateService failed (%d)\n", GetLastError());
		CloseServiceHandle(hSCManager);
		return;
	}
	else
	{
		// Set service description.
		SERVICE_DESCRIPTION sCfgDesc = { SVCDESCRIPTION };
		if (!ChangeServiceConfig2(hService, SERVICE_CONFIG_DESCRIPTION, &sCfgDesc))
		{
			theLog.Warning(MOD_NAME, "Set service description failed");
		}
		// SVCDESCRIPTION
		printf("Service installed successfully\n");
	}

	CloseServiceHandle(hService);
	CloseServiceHandle(hSCManager);
}

//   Deletes a service from the SCM database
VOID  SvcUninstall()
{
	// Get a handle to the SCM database.
	SC_HANDLE hSCManager = OpenSCManager(
		NULL,                    // local computer
		NULL,                    // ServicesActive database 
		SC_MANAGER_ALL_ACCESS);  // full access rights 
	if (NULL == hSCManager)
	{
		printf("OpenSCManager failed (%d)\n", GetLastError());
		return;
	}

	// Get a handle to the service.
	SC_HANDLE hService = OpenService(
		hSCManager,		    // SCM database 
		SVCNAME,            // name of service 
		DELETE);            // need delete access 
	if (NULL == hService)
	{
		printf("OpenService failed (%d)\n", GetLastError());
		CloseServiceHandle(hSCManager);
		return;
	}

	// Delete the service.
	if (!DeleteService(hService))
	{
		printf("DeleteService failed (%d)\n", GetLastError());
	}
	else printf("Service deleted successfully\n");

	CloseServiceHandle(hService);
	CloseServiceHandle(hSCManager);
}

// Read service configuration settings from file.
// FIle is expected to be UNICODE and in the same folder as the 
// ADchangeTracker.exe file.
BOOL ReadConfigFile()
{
	// Get module (exefile) name and path.
	char szFN[MAX_PATH];
	int nLen = GetModuleFileNameA(GetModuleHandle(NULL), szFN, sizeof(szFN));
	if (nLen)
	{	// Replace .exe in string with .cfg
		while (nLen > 0 && szFN[nLen] != '.')	// find last '.' in string.
			nLen--;
		strcpy_s(&szFN[nLen + 1], 4, "cfg");
	}
	else
	{
		theLog.SysErr(MOD_NAME, "Failed to get module name in function ReadConfigFile", "", GetLastError());
		return FALSE;
	}
	HANDLE hFile = CreateFileA(szFN, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (INVALID_HANDLE_VALUE == hFile)
	{
		theLog.SysErr(MOD_NAME, "Failed to open service config file", szFN, GetLastError());
		return FALSE;
	}
	LARGE_INTEGER nFileSize;
	if (!GetFileSizeEx(hFile, &nFileSize))
	{
		theLog.SysErr(MOD_NAME, "Failed to get size of service config file", szFN, GetLastError());
		return FALSE;
	}
	BYTE *pFileData = (BYTE *)malloc(nFileSize.LowPart + 2);
	if (NULL == pFileData)
	{
		theLog.Error(MOD_NAME, "Failed to allocate memory for service config file", szFN);
		return FALSE;
	}
	BOOL fResult = TRUE;
	DWORD dwBytesRead = 0;
	if (!ReadFile(hFile, pFileData, nFileSize.LowPart, &dwBytesRead, NULL))
	{
		theLog.SysErr(MOD_NAME, "Failed to read service config file", szFN, GetLastError());
		fResult = FALSE;
		goto cleanup;
	}
	// Terminate data with 0x00 - needed by _tcstok_s
	*(WORD *)(pFileData + dwBytesRead) = 0;

	ProcessConfigFile(pFileData, dwBytesRead);

cleanup:
	if (pFileData)
		free(pFileData);
	CloseHandle(hFile);
	return fResult;
}

void ProcessConfigFile(BYTE *pFileData, DWORD dwDataLen)
{
	// Check for unicode signature.
	BOOL fIsUnicode = FALSE;
	WORD dwUnicodeSign = *(WORD *)pFileData;
	if (0xFEFF == dwUnicodeSign)
	{
		fIsUnicode = TRUE;
		pFileData += 2;
		dwDataLen -= 2;
	}
	if (!fIsUnicode)
	{
		theLog.Error(MOD_NAME, "Service config file is not UNICODE");
		return;
	}

	TCHAR seps[] = L"\r\n";
	TCHAR* token, *next = 0;
	TCHAR *pszSrc = (TCHAR *)pFileData;

	token = _tcstok_s(pszSrc, seps, &next);
	while (token != NULL)
	{
		ParseConfigFileLine(token);
		token = _tcstok_s(NULL, seps, &next);
	}
}

void ParseConfigFileLine(TCHAR *szLine)
{
	if (szLine[0] == '#')
		return;				// Ignore comment lines.
	TCHAR *next = 0;
	TCHAR *setting = _tcstok_s(szLine, L"=", &next);
	if (!setting)
		return;
	EVENT_PROCESSING_CONFIG &config = theService.GetConfigStruct();
	TCHAR *param = _tcstok_s(NULL, L"\r\n", &next);
	if (_tcsstr(setting, L"SqlConnString") != NULL)
	{
		StringCchCopy(config.szConnectionString, sizeof(config.szConnectionString), param);
	}
	else if (_tcsstr(setting, L"AcceptedEventIDs") != NULL)
	{
		config.nNumElemAcceptedEvts = ParseAcceptedIDs(param,
			config.narrAcceptedEvents, sizeof(config.narrAcceptedEvents) / sizeof(int));
	}
	else if (_tcsstr(setting, L"IgnoredEvents") != NULL)
	{
		config.nNumElemIgnoreEvts = ParseIgnoredEvts(param, 
			config.sarrIgnoreEvts, sizeof(config.sarrIgnoreEvts) / sizeof(IGNORE_EVENTS));
	}
	else if (_tcsstr(setting, L"VerboseLogging") != NULL)
	{
		config.fIsVerboseLogging = ParseBoolParam(param);
	}
}

int ParseAcceptedIDs(TCHAR *szEventIDs, int *pnarrEvents, int nNumElem)
{
	TCHAR szSrc[4096];
	StringCchCopy(szSrc, sizeof(szSrc), szEventIDs);
	TCHAR seps[] = L" ,";
	TCHAR* token, *next = 0;
	int var;
	int elem = 0;

	token = _tcstok_s(szSrc, seps, &next);
	while (token != NULL && elem < nNumElem)
	{
		_stscanf_s(token, L"%d", &var);
		pnarrEvents[elem++] = var;

		token = _tcstok_s(NULL, seps, &next);
	}
	return elem;
}

int ParseIgnoredEvts(TCHAR *szIgnoredEvts, IGNORE_EVENTS *psarrIgnoreEvents, int nNumElem)
{
	TCHAR szSrc[4096];
	StringCchCopy(szSrc, sizeof(szSrc), szIgnoredEvts);
	TCHAR seps[] = L" ,";
	TCHAR* token, *next = 0;
	int elem = 0;

	token = _tcstok_s(szSrc, seps, &next);
	while (token != NULL && elem < nNumElem)
	{
		StringCchCopy(psarrIgnoreEvents[elem].szObjectClass, 
			sizeof(psarrIgnoreEvents[elem].szObjectClass), token);
		++elem;
		token = _tcstok_s(NULL, seps, &next);
	}
	return elem;
}

int ParseBoolParam(TCHAR *szParam)
{
	TCHAR szSrc[64];
	StringCchCopy(szSrc, sizeof(szSrc), szParam);
	_tcslwr_s(szSrc);
	if (_tcsstr(szSrc, L"1") != NULL)
		return TRUE;
	if (_tcsstr(szSrc, L"true") != NULL)
		return TRUE;
	if (_tcsstr(szSrc, L"on") != NULL)
		return TRUE;
	return FALSE;
}

