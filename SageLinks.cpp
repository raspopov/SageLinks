/*
This file is part of SageLinks

Copyright (C) 2015 Nikolay Raspopov <raspopov@cherubicsoft.com>

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

	AfxEnableControlContainer();

	// Create the shell manager, in case the dialog contains
	// any shell tree view or shell list view controls.
	CAutoPtr< CShellManager > pShellManager( new CShellManager );

	// Activate "Windows Native" visual manager for enabling themes in MFC controls
	CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows));

	SetRegistryKey( AFX_IDS_COMPANY_NAME );

	CSageLinksDlg dlg;
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
