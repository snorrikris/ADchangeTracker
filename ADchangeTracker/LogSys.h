#pragma once

class CLogSys
{	// NOTE - only one instance of this object should be instantiated.
public:
	CLogSys(void);
	~CLogSys(void);

private:	// Make assignment operator and copy constructor private to prevent
			// accidental object copy.
	void operator=(CLogSys &source) {  }
	CLogSys(CLogSys &source) {  }

public:
	// Call this before any other member.
	void InitLogSys( HMODULE hModule, int nDaysToKeepOldLogFiles = 30 );

	BOOL IsInitialized() { return m_fInit; }

	// Create new log file in same directory as .exe file.
	// If an log file is open it is closed.
	// Log file will be named as: "exefilename_YYYY-MM-DD_HH-MM-SS.log"
	BOOL CreateNewLogFile();

	void DeleteOldLogFiles( int nOlderThanXdays );

	// Add to log - 'Info' level.
	void Add2LogI( const char *szModule, const char *szLogEvent, 
		const char *szDescription = 0, const char *szNotes = 0 );

	// Add to log - 'Warning' level.
	void Add2LogW( const char *szModule, const char *szLogEvent, 
		const char *szDescription = 0, const char *szNotes = 0 );

	// Add to log - 'Error' level.
	void Add2LogE( const char *szModule, const char *szLogEvent, 
		const char *szDescription = 0, const char *szNotes = 0 );

	// Add to log - custom level.
	void Add2LogLvl( const char *szModule, const char *szLogEvent, const char *szLogLevel,
		const char *szDescription = 0, const char *szNotes = 0 );

	// Add to log - 'Error' level. WIN32 system error code.
	// The 'Notes' field will contain the system error code and description.
	// dwSystemErrorCode = error code from ::GetLastError().
	void Add2LogEsyserr( const char *szModule, const char *szLogEvent, 
		const char *szDescription, DWORD dwSystemErrorCode );

	// By default application name is set to the exe or dll filename.
	void SetApplicationName( const char *szApp );

protected:
	void SaveAndCloseLogFile();
	void WriteToLogFile( const char *szString );
	HANDLE m_hLogFile;
	SYSTEMTIME m_stCurFileStartTime;
	int m_nDaysToKeepOldLogFiles;	// Default is 30 days

	void Init();
	BOOL m_fInit;

	char m_szApplication[128];

	char m_szUserName[128];

	char m_szLoggedOnUsername[128];

	char m_szWorkstationName[128];

	char m_szLogFilenameBase[MAX_PATH]; // Full path and filename base for log files.
									 // Note - "YY-MM-DD_HH-MM-SS.log" is added to name.

	char m_szLogNameBase[MAX_PATH];		// Base name of log files
	char m_szModuleFilename[MAX_PATH];	// Full path and filename of .exe or .dll
	char m_szAppFilename[MAX_PATH];		// Filename of module.
	char m_szModulePath[MAX_PATH];		// Full path to module (not incl. last '\').
	char m_szLogFilePath[MAX_PATH];		// Full path to log files (incl. last '\').

	// Critical section object used to protect log system from being used by 
	// more that one thread at a time.
	CRITICAL_SECTION  m_critsect;

	HMODULE m_hModule;	// handle passed to us when InitLogSys(...) was called.
						// Note - default is = 0.
						// Used when GetModuleFileName called.
};

// Note 'theLogSys' object is instantiated only once and it is global data.
// All threads in 'this' application/service can use it for logging to file.
extern CLogSys theLogSys;

