//-----------------------------------------------------------------------
// 
//  Copyright (C) 2006 Microsoft Corporation.  All rights reserved.
// 
// THIS CODE AND INFORMATION ARE PROVIDED AS IS WITHOUT WARRANTY OF ANY
// KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
// IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//-----------------------------------------------------------------------

#define WIN32_LEAN_AND_MEAN

// Windows header files.
#include <windows.h>
#include <msi.h>

// C runtime header files.
#include <stdlib.h>
#include <tchar.h>

// Functions used from msi.dll.
typedef INSTALLSTATE (WINAPI *FMsiLocateComponent) (LPCTSTR, LPTSTR, DWORD*);
#ifdef UNICODE
    #define FUNCNAME_MsiLocateComponent "MsiLocateComponentW"
#else
    #define FUNCNAME_MsiLocateComponent "MsiLocateComponentA"
#endif

// Changed By Brian Randell
// Check for an Office 2003 or Office 2007 Installation
// Returns 0 if the Office MSO Component specified at the command line is installed.
// For example, use the following command lines to check for a certain product:
// - ComponentCheck.exe {3EC1EAE0-A256-411D-B00B-016CA8376078}' // Microsoft Office 2003
// - ComponentCheck.exe {0638C49D-BB8B-4CD1-B191-050E8F325736}' // Microsoft Office 2007
int APIENTRY _tWinMain(
    HINSTANCE hInstance,
    HINSTANCE hPrevInstance,
    LPTSTR    lpCmdLine,
    int       nCmdShow)
{
    UINT errCode = 0;
    HMODULE hMsiLib = LoadLibrary(_T("MSI.DLL"));
    if (hMsiLib == NULL)
    {
        errCode = GetLastError();
        if (errCode == 0) // make sure the error code is not 0
            errCode = ERROR_MOD_NOT_FOUND;
    }
    else
    {
        FMsiLocateComponent pfnMsiLocateComponent = (FMsiLocateComponent)GetProcAddress(hMsiLib, FUNCNAME_MsiLocateComponent);
        if (pfnMsiLocateComponent == NULL)
        {
            errCode = GetLastError();
            if (errCode == 0) // make sure the error code is not 0
                errCode = ERROR_PROC_NOT_FOUND;
        }
        else
        {
            // Loop through the specified components in the arguments.
            for (int i = 1; i < __argc; i++)
            {
                LPCTSTR szComponentCode = __targv[i];
                INSTALLSTATE state = (*pfnMsiLocateComponent)(szComponentCode, NULL, 0);
                if (state != INSTALLSTATE_LOCAL)
                {
                    errCode = ERROR_UNKNOWN_COMPONENT;
                    break;
                }
            }
        }
    }

    if (hMsiLib != NULL)
    {
        FreeLibrary(hMsiLib);
    }
    return errCode;
}

