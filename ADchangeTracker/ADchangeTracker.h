#pragma once
#include "EventProcessing.h"

VOID SvcInstall();
VOID SvcUninstall();

// Read config file functions.
BOOL ReadConfigFile();
void ProcessConfigFile(BYTE *pFileData, DWORD dwDataLen);
void ParseConfigFileLine(TCHAR *szLine);
int ParseAcceptedIDs(TCHAR *szEventIDs, int *pnarrEvents, int nNumElem);
int ParseIgnoredEvts(TCHAR *szIgnoredEvts, IGNORE_EVENTS *psarrIgnoreEvents, int nNumElem);
BOOL ParseBoolParam(TCHAR *szParam);
