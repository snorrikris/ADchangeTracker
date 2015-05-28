#pragma once

#include "targetver.h"

#include <stdio.h>
#include <conio.h>
#include <tchar.h>
#include <windows.h>
#include <winevt.h>
#include <windows.h>
#include <Strsafe.h>
#include "pugixml.hpp"
#pragma comment(lib, "wevtapi.lib")

#include <assert.h>

// SQL stuff
#define _WIN32_DCOM
//#import "C:\Program Files (x86)\Common Files\System\ado\msado60_Backcompat.tlb" \
//    no_namespace rename("EOF", "EndOfFile")
#import "C:\Program Files\Common Files\System\ADO\msado15.dll" \
    no_namespace rename("EOF", "EndOfFile")
#include <ole2.h>
#include <oledb.h>
inline void TESTHR(HRESULT x) { if FAILED(x) _com_issue_error(x); };
