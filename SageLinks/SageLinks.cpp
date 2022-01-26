// This is an open source non-commercial project. Dear PVS-Studio, please check it.
// PVS-Studio Static Code Analyzer for C, C++ and C#: http://www.viva64.com
/*
This file is part of SageLinks

Copyright (C) 2015-2020 Nikolay Raspopov <raspopov@cherubicsoft.com>

This program is free software : you can redistribute it and / or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
( at your option ) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.If not, see < http://www.gnu.org/licenses/>.
*/

#include "stdafx.h"
#include "SageLinks.h"
#include "SageLinksDlg.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#define new DEBUG_NEW
#endif


// CSageLinksApp

BEGIN_MESSAGE_MAP(CSageLinksApp, CWinApp)
END_MESSAGE_MAP()

// CSageLinksApp construction

CSageLinksApp::CSageLinksApp()
{
	m_dwRestartManagerSupportFlags = AFX_RESTART_MANAGER_SUPPORT_RESTART;
}

// The one and only CSageLinksApp object

CSageLinksApp theApp;

// CSageLinksApp initialization

BOOL CSageLinksApp::InitInstance()
{
	const INITCOMMONCONTROLSEX InitCtrls = { sizeof( InitCtrls ), ICC_WIN95_CLASSES };
	InitCommonControlsEx( &InitCtrls );

	CWinApp::InitInstance();

	AfxOleInit();

	AfxEnableControlContainer();

	// Create the shell manager, in case the dialog contains any shell tree view or shell list view controls.
	CAutoPtr< CShellManager > pShellManager( new CShellManager );

	// Activate "Windows Native" visual manager for enabling themes in MFC controls
	CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows));

	SetRegistryKey( AFX_IDS_COMPANY_NAME );

	// Parse initial folder
	CCommandLineInfo cmdInfo;
	ParseCommandLine( cmdInfo );

	CSageLinksDlg dlg;
	if ( ! cmdInfo.m_strFileName.IsEmpty() )
	{
		dlg.m_sPath = cmdInfo.m_strFileName.Trim( _T(" \"") );
	}

	m_pMainWnd = &dlg;
	dlg.DoModal();

	// Since the dialog has been closed, return FALSE so that we exit the
	//  application, rather than start the application's message pump.
	return FALSE;
}

BOOL CSageLinksApp::ProcessMessageFilter(int code, LPMSG lpMsg)
{
	if ( lpMsg->hwnd == m_pMainWnd->m_hWnd || ::IsChild( m_pMainWnd->m_hWnd, lpMsg->hwnd ) )
	{
		if ( lpMsg->message == WM_KEYDOWN )
		{
			// Emulate key down message for dialog
			if ( ( (CSageLinksDlg*)m_pMainWnd )->OnKeyDown( (UINT)lpMsg->wParam, (UINT)( lpMsg->lParam & 0xffff ), (UINT)( ( lpMsg->lParam >> 16 ) & 0xffff ) ) )
				return TRUE;
		}
	}

	return CWinApp::ProcessMessageFilter( code, lpMsg );
}

CString ErrorMessage(HRESULT hr)
{
	static const LPCTSTR szModules [] =
	{
		_T("user32.dll"),
		_T("netapi32.dll"),
		_T("netmsg.dll"),
		_T("netevent.dll"),
		_T("spmsg.dll"),
		_T("wininet.dll"),
		_T("ntdll.dll"),
		_T("ntdsbmsg.dll"),
		_T("mprmsg.dll"),
		_T("IoLogMsg.dll"),
		_T("NTMSEVT.DLL"),
		_T("ws2_32.dll")
	};

	CString sError;

	if ( hr == ERROR_EXTENDED_ERROR )
	{
		CString sDescription, sProvider;
		DWORD err = hr;
		DWORD result = WNetGetLastError( &err, sDescription.GetBuffer( 1024 ), 1024, sProvider.GetBuffer( 256 ), 256 );
		sDescription.ReleaseBuffer();
		sProvider.ReleaseBuffer();
		if ( NO_ERROR == result )
		{
			sError = sDescription;
		}
	}

	if ( sError.IsEmpty() )
	{
		LPTSTR lpszTemp = nullptr;
		FormatMessage( FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM, nullptr, hr, 0, (LPTSTR)&lpszTemp, 0, nullptr );
		for ( int i = 0; ! lpszTemp && i < _countof( szModules ); ++i )
		{
			if ( HMODULE hModule = LoadLibraryEx( szModules[ i ], nullptr, LOAD_LIBRARY_AS_DATAFILE ) )
			{
				FormatMessage( FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_HMODULE, hModule, hr, 0, (LPTSTR)&lpszTemp, 0, nullptr );
				FreeLibrary( hModule );
			}
		}

		if ( lpszTemp )
		{
			sError = lpszTemp;
			LocalFree( lpszTemp );
		}
	}

	if ( sError.IsEmpty() )
	{
		sError.LoadString( IDS_UNKNOWN_ERROR );
	}

	sError.TrimRight( _T(". \r\n") );
	sError.AppendFormat( _T(" (0x%08x)"), hr );
	return sError;
}
