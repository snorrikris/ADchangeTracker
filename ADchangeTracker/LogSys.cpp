#include "StdAfx.h"
#include "LogSys.h"
#include <time.h>
#include <stdio.h>

/////////////////////////////////////////////////////////////////////////////////////
// The one and only CLogSys object
CLogSys theLog;
/////////////////////////////////////////////////////////////////////////////////////

CLogSys::CLogSys(void)
{
	// Initialize the critical section object
	::InitializeCriticalSection( &m_critsect );
	m_hLogFile = 0;
	m_fInit = FALSE;
	m_hModule = 0;
	m_nDaysToKeepOldLogFiles = 30;
	memset(&m_stCurFileStartTime, 0, sizeof(SYSTEMTIME));
}

void CLogSys::InitLogSys( HMODULE hModule, int nDaysToKeepOldLogFiles )
{
	m_hModule = hModule;
	m_nDaysToKeepOldLogFiles = nDaysToKeepOldLogFiles;
	Init();
}

CLogSys::~CLogSys(void)
{
	SaveAndCloseLogFile();
	::DeleteCriticalSection( &m_critsect );
}

// Add to log - 'Info' level.
void CLogSys::Info( const char *szModule, const char *szLogEvent, 
						 const char *szDescription /*= 0*/,	const char *szNotes/*= 0*/ )
{
	Add2LogLvl( szModule, szLogEvent, "Info", szDescription, szNotes );
}

// Add to log - 'Warning' level.
void CLogSys::Warning( const char *szModule, const char *szLogEvent,
	const char *szDescription/*= 0*/, const char *szNotes/*= 0*/ )
{
	Add2LogLvl( szModule, szLogEvent, "Warning", szDescription, szNotes );
}

// Add to log - 'Error' level.
void CLogSys::Error( const char *szModule, const char *szLogEvent, 
	const char *szDescription/*= 0*/, const char *szNotes/*= 0*/ )
{
	Add2LogLvl( szModule, szLogEvent, "Error", szDescription, szNotes );
}

void CLogSys::Add2LogLvl( const char *szModule, const char *szLogEvent, 
	const char *szLogLevel, const char *szDescription /*= 0*/, 
	const char *szNotes /*= 0*/ )
{
	if( !m_hLogFile )
		CreateNewLogFile();

	::EnterCriticalSection( &m_critsect );

	// Detect if current log file was created yesterday.
	SYSTEMTIME now, &cur = m_stCurFileStartTime;
	GetSystemTime(&now);
	if (now.wYear != cur.wYear || now.wMonth != cur.wMonth || now.wDay != cur.wDay)
	{
		CreateNewLogFile();	// Create new log file every day.
	}

	DWORD dwThreadID = ::GetCurrentThreadId();
	char szThreadID[20];
	sprintf_s( szThreadID, sizeof(szThreadID), "%u", dwThreadID );
	WriteToLogFile( szThreadID );
	WriteToLogFile( "\t" );

	char szTM[128];
	struct tm newtime;
	time_t aclock;
	time( &aclock );   // Get time in seconds
	localtime_s( &newtime, &aclock );   // Convert time to struct tm form 
	strftime( szTM, 128, "%Y-%m-%d %H:%M:%S", &newtime );
	WriteToLogFile( szTM );
	WriteToLogFile( "\t" );
	WriteToLogFile( m_szWorkstationName );
	WriteToLogFile( "\t" );
	WriteToLogFile( m_szUserName );
	WriteToLogFile( "\t" );
	WriteToLogFile( szModule );
	WriteToLogFile( "\t" );
	WriteToLogFile( szLogLevel );
	WriteToLogFile( "\t" );
	WriteToLogFile( szLogEvent );
	WriteToLogFile( "\t" );
	WriteToLogFile( szDescription );
	WriteToLogFile( "\t" );
	WriteToLogFile( szNotes );
	WriteToLogFile( "\r\n" );

	::LeaveCriticalSection( &m_critsect );
}

void CLogSys::SysErr( const char *szModule, const char *szLogEvent, 
	const char *szDescription, DWORD dwSystemErrorCode )
{
	LPVOID lpMsgBuf = 0;
	DWORD dwStrLen = ::FormatMessageA( FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
		NULL, dwSystemErrorCode, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
		(LPSTR)&lpMsgBuf, 0, NULL );
	if( dwStrLen )
	{	// Trim CRLF from string end.
		char *pc = (char *)lpMsgBuf;
		pc[dwStrLen - 2] = 0;
	}

	char sz[1024];
	if( lpMsgBuf )
		sprintf_s( sz, sizeof(sz), "Error code: 0x%.8X - %s", dwSystemErrorCode, lpMsgBuf );
	else
		sprintf_s( sz, sizeof(sz), "Error code: 0x%.8X", dwSystemErrorCode );
	Error( szModule, szLogEvent, szDescription, sz );

	::LocalFree(lpMsgBuf);
}

void CLogSys::SetApplicationName( const char *szApp )
{
	::EnterCriticalSection( &m_critsect );

	strcpy_s( m_szApplication, sizeof(m_szApplication), szApp );

	::LeaveCriticalSection( &m_critsect );
}

BOOL CLogSys::CreateNewLogFile()
{
	::EnterCriticalSection( &m_critsect );

	if( !m_fInit )
		Init();

	DeleteOldLogFiles(m_nDaysToKeepOldLogFiles);	// Delete Log files older than X days

	BOOL fRetVal = TRUE;

	if( m_hLogFile )
		SaveAndCloseLogFile();

	// Log file will be named as: "exefilename_YYYY-MM-DD_HH-MM-SS.log"
	char szTM[128];
	struct tm newtime;
	time_t aclock;
	time( &aclock );   // Get time in seconds
	localtime_s( &newtime, &aclock );   // Convert time to struct tm form 
	strftime( szTM, 128, "%Y-%m-%d_%H-%M-%S", &newtime );

	char szFileName[MAX_PATH];
	strcpy_s( szFileName, sizeof(szFileName), m_szLogFilenameBase );
	strcat_s(szFileName, sizeof(szFileName), szTM);
	strcat_s(szFileName, sizeof(szFileName), ".log");

	m_hLogFile = ::CreateFileA( szFileName, GENERIC_WRITE, FILE_SHARE_READ, NULL,
		CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL );
	if( m_hLogFile == INVALID_HANDLE_VALUE )
	{
		m_hLogFile = 0;
		memset(&m_stCurFileStartTime, 0, sizeof(SYSTEMTIME));
		fRetVal = FALSE;
	}
	else
	{	
		GetSystemTime(&m_stCurFileStartTime);	// Save create time for current file.

		// Write header to log file.
		WriteToLogFile( 
			"ThreadID\tLogDate\tWorkStation\tUserName\tModule\tLogLevel\tLogEvent\tDescription\tNotes\r\n" );
	}

	::LeaveCriticalSection( &m_critsect );
	return fRetVal;
}

void CLogSys::DeleteOldLogFiles( int nOlderThanXdays )
{
	char szDelFile[MAX_PATH];
	char szDirSrch[MAX_PATH];
	strcpy_s( szDirSrch, sizeof(szDirSrch), m_szLogFilenameBase );
	strcat_s( szDirSrch, sizeof(szDirSrch), "*.log" );

	// Calculate 'now' - 30 days as FILETIME data structure.
	SYSTEMTIME stCurSysTime;
	::GetSystemTime( &stCurSysTime );
	FILETIME stFT;
	::SystemTimeToFileTime( &stCurSysTime, &stFT );
	ULARGE_INTEGER *pu64 = (ULARGE_INTEGER *)&stFT;
	// Note - FILETIME is equal to ULARGE_INTEGER structure.

	// X days at FILETIME res.100nS                 
	unsigned __int64 u64_FT_days 
		= (unsigned __int64)10000000			// 1SEC
		* (unsigned __int64)60					// MM
		* (unsigned __int64)60					// HH
		* (unsigned __int64)24					// DD
		* (unsigned __int64)nOlderThanXdays;	// Days.
	pu64->QuadPart -= u64_FT_days;
	// Impl. note - to specify a 64bit constant the 'i64' integer suffix can be used.
	// But - '10000000i64' just seems ugly :)

	WIN32_FIND_DATAA stFindFileData;
	HANDLE hFileFind = ::FindFirstFileA( szDirSrch, &stFindFileData );
	if( hFileFind != INVALID_HANDLE_VALUE )
	{
		do
		{	// Is file older then X days.
			if( ::CompareFileTime( &stFindFileData.ftCreationTime, &stFT ) < 0 )
			{
				strcpy_s( szDelFile, sizeof(szDelFile), m_szLogFilePath );
				strcat_s( szDelFile, sizeof(szDelFile), stFindFileData.cFileName );
				::DeleteFileA( szDelFile );
			}
		} while( ::FindNextFileA( hFileFind, &stFindFileData ) != 0 );
		::FindClose( hFileFind );
	}
}

void CLogSys::SaveAndCloseLogFile()
{
	::EnterCriticalSection( &m_critsect );

	if( m_hLogFile )
		::CloseHandle( m_hLogFile );
	m_hLogFile = 0;

	::LeaveCriticalSection( &m_critsect );
}

void CLogSys::WriteToLogFile( const char *szString )
{
	if( !szString )
		return;

	::EnterCriticalSection( &m_critsect );

	if( m_hLogFile )
	{
		DWORD dwBytesWritten;
		::WriteFile( m_hLogFile, szString, strlen( szString ), &dwBytesWritten, NULL );
	}

	::LeaveCriticalSection( &m_critsect );
}

void CLogSys::Init()
{
	if( m_fInit )
		return;

	// Get workstation name.
	DWORD dw = 128;
	::GetComputerNameA( m_szWorkstationName, &dw );

	// Get 'logged on' user name.
	dw = 128;
	::GetUserNameA( m_szLoggedOnUsername, &dw );

	// Get module (exefile) name and path.
	int nLen, nBufLen = MAX_PATH;
	char *szFN = m_szModuleFilename;
	nLen = ::GetModuleFileNameA( m_hModule, szFN, nBufLen );

	strcpy_s( m_szLogFilenameBase, sizeof(m_szLogFilenameBase), m_szModuleFilename );
	szFN = m_szLogFilenameBase;
	if( nLen )
	{	// Remove .exe (or .dll) from string
		while( nLen > 0 && szFN[nLen] != '.' )
			nLen--;
		szFN[nLen] = '_';	// replace '.' with '_'
		szFN[nLen+1] = 0;
	}

	strcpy_s( m_szModulePath, sizeof(m_szModulePath), m_szModuleFilename );
	nLen = strlen( m_szModulePath );
	szFN = m_szModulePath;
	if( nLen )
	{	// find last '\' in string and replace with 0
		while( nLen > 0 && szFN[nLen] != '\\' )
			nLen--;
		szFN[nLen] = 0;
	}

	strcpy_s( m_szLogNameBase, sizeof(m_szLogNameBase), &m_szLogFilenameBase[nLen + 1] );
	strcpy_s( m_szAppFilename, sizeof(m_szAppFilename), &m_szModuleFilename[nLen + 1] );

	strcpy_s( m_szLogFilePath, sizeof(m_szLogFilePath), m_szModulePath );
	strcat_s( m_szLogFilePath, sizeof(m_szLogFilePath), "\\" );

	// Put log files in "Logs" directory if it excists.
	char szDirSrch[MAX_PATH];
	strcpy_s(szDirSrch, sizeof(szDirSrch), m_szModulePath);
	strcat_s(szDirSrch, sizeof(szDirSrch), "\\Logs");
	WIN32_FIND_DATAA stFindFileData;
	HANDLE hFind = ::FindFirstFileA( szDirSrch, &stFindFileData );
	if( hFind != INVALID_HANDLE_VALUE )
	{
		strcpy_s(m_szLogFilenameBase, sizeof(m_szLogFilenameBase), m_szModulePath);
		strcat_s(m_szLogFilenameBase, sizeof(m_szLogFilenameBase), "\\Logs\\");
		strcat_s(m_szLogFilenameBase, sizeof(m_szLogFilenameBase), m_szLogNameBase);
		::FindClose( hFind );
		strcat_s(m_szLogFilePath, sizeof(m_szLogFilePath), "Logs\\");
	}

	strcpy_s(m_szUserName, sizeof(m_szUserName), m_szLoggedOnUsername);
	strcpy_s(m_szApplication, sizeof(m_szApplication), m_szAppFilename);
	nLen = strlen(m_szApplication);
	if (nLen)
	{	// Remove .exe from string
		while (nLen > 0 && m_szApplication[nLen] != '.')
			nLen--;
		m_szApplication[nLen] = 0;
	}

	// Get path to %PROGRAMDATA% folder.
	PWSTR pszPath = 0;
	HRESULT hr = SHGetKnownFolderPath(FOLDERID_ProgramData, 0, NULL, &pszPath);
	if (hr == S_OK)
	{
		char szLogDir[MAX_PATH];
		memset(szLogDir, 0, sizeof(szLogDir));
		nLen = _tcslen(pszPath);
		CharToOemBuff(pszPath, szLogDir, nLen);
		strcat_s(szLogDir, sizeof(szLogDir), "\\");
		strcat_s(szLogDir, sizeof(szLogDir), m_szApplication);
		BOOL fCreateResult = CreateDirectoryA(szLogDir, NULL);
		DWORD dwError = GetLastError();
		if (fCreateResult || ERROR_ALREADY_EXISTS == dwError)
		{
			// Directory was created or already existed.
			strcpy_s(m_szLogFilePath, sizeof(m_szLogFilePath), szLogDir);
			strcat_s(m_szLogFilePath, sizeof(m_szLogFilePath), "\\");

			strcpy_s(m_szLogFilenameBase, sizeof(m_szLogFilenameBase), m_szLogFilePath);
			strcat_s(m_szLogFilenameBase, sizeof(m_szLogFilenameBase), m_szApplication);
			strcat_s(m_szLogFilenameBase, sizeof(m_szLogFilenameBase), "_");
		}
	}
	CoTaskMemFree(static_cast<void*>(pszPath));
	pszPath = 0;

	m_fInit = TRUE;
}

