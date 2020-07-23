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
#include "ShellMenu.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#define new DEBUG_NEW
#endif

#define CX_ICON		20	// pixels

// Column numbers

#define COL_SOURCE	0
#define COL_TYPE	1
#define COL_TARGET	2
#define COL_RESULT	3

// CSageLinksDlg dialog

CSageLinksDlg::CSageLinksDlg(CWnd* pParent /*=NULL*/)
	: CDialogExSized	( IDD_SAGELINKS_DIALOG, pParent )
	, m_sPath	( theApp.GetProfileString( OPT_SECTION, OPT_PATH, _T("C:\\") ) )
	, m_hIcon	( AfxGetApp()->LoadIcon( IDR_MAINFRAME ) )
	, m_hThread	( NULL )
	, m_pFlag	( FALSE, TRUE )
	, m_nTotal	( 0 )
	, m_nBad	( 0 )
	, m_bSort	( FALSE )
	, m_nImageSuccess	( 0 )
	, m_nImageError		( 0 )
	, m_nImageUnknown	( 0 )
	, m_nImageSymbolic	( 0 )
	, m_nImageJunction	( 0 )
	, m_nImageShortcut	( 0 )
{
}

void CSageLinksDlg::DoDataExchange(CDataExchange* pDX)
{
	__super::DoDataExchange( pDX );

	DDX_Control( pDX, IDC_LIST, m_wndList );
	DDX_Control( pDX, IDC_STATUS, m_wndStatus );
	DDX_Control( pDX, IDC_BROWSE, m_wndBrowse );
}

IMPLEMENT_DYNAMIC(CSageLinksDlg, CDialogExSized)

BEGIN_MESSAGE_MAP(CSageLinksDlg, CDialogExSized)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_TIMER()
	ON_WM_DESTROY()
	ON_WM_SIZE()
	ON_WM_GETMINMAXINFO()
	ON_NOTIFY( NM_DBLCLK, IDC_LIST, &CSageLinksDlg::OnNMDblclkList )
	ON_NOTIFY( NM_RCLICK, IDC_LIST, &CSageLinksDlg::OnNMRClickList )
	ON_NOTIFY( HDN_ITEMCLICK, 0, &CSageLinksDlg::OnHdnItemclickList )
END_MESSAGE_MAP()

// CSageLinksDlg message handlers

BOOL CSageLinksDlg::OnInitDialog()
{
	__super::OnInitDialog();

	SetIcon( m_hIcon, TRUE );		// Set big icon
	SetIcon( m_hIcon, FALSE );		// Set small icon

	GetWindowRect( m_rcInitial );

	m_oImages.Create( 16, 16, ILC_COLOR32 | ILC_MASK, 0, 100 );
	m_nImageSuccess = m_oImages.Add( (HICON)LoadImage( AfxGetResourceHandle(), MAKEINTRESOURCE( IDI_RESULT_SUCCESS ), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR | LR_SHARED ) );
	m_nImageError = m_oImages.Add( (HICON)LoadImage( AfxGetResourceHandle(), MAKEINTRESOURCE( IDI_RESULT_ERROR ), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR | LR_SHARED ) );
	m_nImageUnknown = m_oImages.Add( (HICON)LoadImage( AfxGetResourceHandle(), MAKEINTRESOURCE( IDI_UNKNOWN ), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR | LR_SHARED ) );
	m_nImageSymbolic = m_oImages.Add( (HICON)LoadImage( AfxGetResourceHandle(), MAKEINTRESOURCE( IDI_SYMBOLIC ), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR | LR_SHARED ) );
	m_nImageJunction = m_oImages.Add( (HICON)LoadImage( AfxGetResourceHandle(), MAKEINTRESOURCE( IDI_JUNCTION ), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR | LR_SHARED ) );
	m_nImageShortcut = m_oImages.Add( (HICON)LoadImage( AfxGetResourceHandle(), MAKEINTRESOURCE( IDI_SHORTCUT ), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR | LR_SHARED ) );

	m_wndList.SetImageList( &m_oImages, LVSIL_SMALL );
	m_wndList.SetExtendedStyle( m_wndList.GetExtendedStyle() | LVS_EX_DOUBLEBUFFER | LVS_EX_FULLROWSELECT | LVS_EX_LABELTIP | LVS_EX_SUBITEMIMAGES );

	CRect rc;
	m_wndList.GetClientRect( &rc );
	const int width = ( rc.Width() - CX_ICON - CX_ICON - GetSystemMetrics( SM_CXVSCROLL ) - 5 ) / 2;
	m_wndList.InsertColumn( COL_SOURCE, _T( "Source" ), LVCFMT_LEFT, width );
	m_wndList.InsertColumn( COL_TYPE, _T( "" ), LVCFMT_CENTER, CX_ICON );
	m_wndList.InsertColumn( COL_TARGET, _T( "Target" ), LVCFMT_LEFT, width );
	m_wndList.InsertColumn( COL_RESULT, _T( "" ), LVCFMT_CENTER, CX_ICON );

	HDITEM it = { HDI_FORMAT };
	m_wndList.GetHeaderCtrl()->GetItem( COL_RESULT, &it );
	it.fmt |= HDF_SORTDOWN;
	m_wndList.GetHeaderCtrl()->SetItem( COL_RESULT, &it );

	m_wndBrowse.SetWindowText( m_sPath );

	RestoreWindowPlacement();

	SetTimer( 100, 200, NULL );

	Start();

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CSageLinksDlg::OnDestroy()
{
	KillTimer( 100 );

	Stop();

	ClearList();

	CString sPath;
	m_wndBrowse.GetWindowText( sPath );
	theApp.WriteProfileString( OPT_SECTION, OPT_PATH, sPath );

	__super::OnDestroy();
}

void CSageLinksDlg::Stop()
{
	if ( m_hThread )
	{
		CWaitCursor wc;

		m_pFlag.SetEvent();

		WaitForSingleObject( m_hThread, INFINITE );

		CloseHandle( m_hThread );
		m_hThread = NULL;
	}
}

void CSageLinksDlg::Start()
{
	if ( ! m_hThread )
	{
		CWaitCursor wc;

		CSingleLock oLock( &m_pSection, TRUE );

		CString sOk;
		sOk.LoadString( IDS_STOP );
		GetDlgItem( IDOK )->SetWindowText( sOk );

		m_wndStatus.SetWindowText( _T( "" ) );

		ClearList();

		CString sRoot;
		m_wndBrowse.GetWindowText( sRoot );
		m_oDirs.AddTail( sRoot );

		m_pFlag.ResetEvent();

		unsigned id;
		m_hThread = (HANDLE)_beginthreadex( NULL, 0, &CSageLinksDlg::ThreadStub, this, 0, &id );
	}
}

unsigned __stdcall CSageLinksDlg::ThreadStub(void* param)
{
	CSageLinksDlg* pThis = (CSageLinksDlg*)param;
	pThis->Thread();
	return 0;
}

void CSageLinksDlg::OnNewItem(CLink* pLink)
{
	if ( pLink->m_nType != LinkType::Unknown )
	{
		const BOOL bIsLastVisible = ( m_nTotal == 0 ) || (BOOL)ListView_IsItemVisible( m_wndList.GetSafeHwnd(), m_nTotal - 1 );

		const LVITEM itSource = { LVIF_IMAGE | LVIF_TEXT | LVIF_PARAM, m_nTotal, COL_SOURCE, 0, 0, (LPTSTR)(LPCTSTR)pLink->m_sSource, 0,
			( pLink->m_hIcon ? m_oImages.Add( pLink->m_hIcon ) : m_nImageUnknown ), (LPARAM)pLink };
		const int index = m_wndList.InsertItem( &itSource );
		++ m_nTotal;
		if ( ! pLink->m_bResult ) ++ m_nBad;

		const LVITEM itType = { LVIF_IMAGE, index, COL_TYPE, 0, 0, NULL, 0,
			( ( pLink->m_nType == LinkType::Symbolic ) ? m_nImageSymbolic :
			( ( pLink->m_nType == LinkType::Junction ) ? m_nImageJunction :
			( ( pLink->m_nType == LinkType::Shortcut ) ? m_nImageShortcut : m_nImageUnknown ) ) ) };
		m_wndList.SetItem( &itType );

		const LVITEM itTarget = { LVIF_TEXT, index, COL_TARGET, 0, 0, (LPTSTR)(LPCTSTR)pLink->m_sTarget };
		m_wndList.SetItem( &itTarget );

		const LVITEM itResult = { LVIF_IMAGE, index, COL_RESULT, 0, 0, NULL, 0, pLink->m_bResult ? m_nImageSuccess : m_nImageError };
		m_wndList.SetItem( &itResult );

		m_bSort = TRUE;

		if ( bIsLastVisible )
			m_wndList.EnsureVisible( m_nTotal - 1, FALSE );
	}
	else
	{
		delete pLink;

		// Done
		CString sOk;
		sOk.LoadString( IDS_START );
		GetDlgItem( IDOK )->SetWindowText( sOk );

		CString sText;
		sText.Format( IDS_DONE, m_nTotal, m_nBad );
		m_wndStatus.SetWindowText( sText );

		// Clean-up
		Stop();
	}
}

void CSageLinksDlg::OnTimer( UINT_PTR nIDEvent )
{
	CString sStatus;
	{
		CSingleLock oLock( &m_pSection, TRUE );

		while ( ! m_pIncoming.IsEmpty() )
		{
			OnNewItem( m_pIncoming.RemoveHead() );
		}

		sStatus = m_sStatus;
	}
	if ( sStatus != m_sOldStatus )
	{
		m_sOldStatus = sStatus;

		if ( ! sStatus.IsEmpty() )
		{
			CString sText;
			sText.Format( IDS_WORK, (LPCTSTR)sStatus );
			m_wndStatus.SetWindowText( sText );
		}
	}

	if ( m_bSort )
	{
		m_bSort = FALSE;
		SortList();
	}

	__super::OnTimer( nIDEvent );
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CSageLinksDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		__super::OnPaint();
	}
}

HCURSOR CSageLinksDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CSageLinksDlg::OnOK()
{
	CWnd* pFocus = GetFocus();
	GetDlgItem( IDOK )->EnableWindow( FALSE );

	if ( m_hThread )
	{
		Stop();
	}
	else
	{
		Start();
	}

	GetDlgItem( IDOK )->EnableWindow( TRUE );
	pFocus->SetFocus();
}

void CSageLinksDlg::OnSize( UINT nType, int cx, int cy )
{
	if ( m_wndList.m_hWnd )
	{
		CRect rc;
		m_wndList.GetClientRect( &rc );
		const int width = ( rc.Width() - CX_ICON - CX_ICON - 5 ) / 2;
		m_wndList.SetColumnWidth( COL_SOURCE, width );
		m_wndList.SetColumnWidth( COL_TYPE, CX_ICON );
		m_wndList.SetColumnWidth( COL_TARGET, width );
		m_wndList.SetColumnWidth( COL_RESULT, CX_ICON );
	}

	__super::OnSize( nType, cx, cy );
}

void CSageLinksDlg::OnGetMinMaxInfo(MINMAXINFO* lpMMI)
{
	__super::OnGetMinMaxInfo( lpMMI );

	lpMMI->ptMinTrackSize.x = m_rcInitial.Width();
	lpMMI->ptMinTrackSize.y = m_rcInitial.Height();
}

void CSageLinksDlg::OnNMDblclkList(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast< LPNMITEMACTIVATE >( pNMHDR );

	if ( pNMItemActivate->iItem >= 0 )
	{
		CWaitCursor wc;

		const CLink* pLink = (const CLink*)m_wndList.GetItemData( pNMItemActivate->iItem );
		ShellExecute( GetSafeHwnd(), NULL, _T("explorer.exe"), _T("/select, \"") + pLink->m_sSource + _T("\""), NULL, SW_SHOWNORMAL );
	}

	*pResult = 0;
}

void CSageLinksDlg::OnNMRClickList( NMHDR *pNMHDR, LRESULT *pResult )
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>( pNMHDR );

	if ( pNMItemActivate->iItem >= 0 )
	{
		CWaitCursor wc;

		CStringList oFiles;
		for ( int i = 0; i < m_nTotal; ++i )
		{
			if (  m_wndList.GetItemState( i, LVIS_SELECTED ) == LVIS_SELECTED )
			{
				const CLink* pLink = (const CLink*)m_wndList.GetItemData( i );
				oFiles.AddHead( pLink->m_sSource );
			}
		}

		if ( ! oFiles.IsEmpty() )
		{
			CMenu oMenu;
			if ( oMenu.CreatePopupMenu() )
			{
				CPoint ptAction( pNMItemActivate->ptAction );
				m_wndList.ClientToScreen( &ptAction );

				DoExplorerMenu( GetSafeHwnd(), oFiles, ptAction, oMenu.GetSafeHmenu() );

				// Re-check list for deleted items
				for ( int i = m_nTotal - 1; i >= 0; -- i )
				{
					if ( m_wndList.GetItemState( i, LVIS_SELECTED ) == LVIS_SELECTED )
					{
						CLink* pLink = (CLink*)m_wndList.GetItemData( i );
						if ( GetFileAttributes( pLink->m_sSource ) == INVALID_FILE_ATTRIBUTES )
						{
							VERIFY( m_wndList.DeleteItem( i ) );

							-- m_nTotal;
							if ( ! pLink->m_bResult )
								-- m_nBad;

							delete pLink;
						}
					}
				}
			}
		}
	}

	*pResult = 0;
}

void CSageLinksDlg::OnHdnItemclickList( NMHDR *pNMHDR, LRESULT *pResult )
{
	NMLISTVIEW* pLV = (NMLISTVIEW*)pNMHDR;

	CHeaderCtrl* pHeader = m_wndList.GetHeaderCtrl();
	const int count = pHeader->GetItemCount();
	HDITEM it = { HDI_FORMAT };
	for ( int i = 0; i < count; ++i )
	{
		pHeader->GetItem( i, &it );
		if ( i == pLV->iItem )
			it.fmt = ( it.fmt & ~( HDF_SORTDOWN | HDF_SORTUP ) ) | ( ( it.fmt & HDF_SORTUP ) ? HDF_SORTDOWN : HDF_SORTUP );
		else
			it.fmt &= ~( HDF_SORTDOWN | HDF_SORTUP );
		pHeader->SetItem( i, &it );
	}

	SortList();

	*pResult = 0;
}

void CSageLinksDlg::SortList()
{
	const CHeaderCtrl* pHeader = m_wndList.GetHeaderCtrl();
	const int count = pHeader->GetItemCount();
	HDITEM it = { HDI_FORMAT };
	for ( int i = 0; i < count; ++i )
	{
		pHeader->GetItem( i, &it );
		if ( it.fmt & ( HDF_SORTUP | HDF_SORTDOWN ) )
		{
			m_wndList.SortItems( &CSageLinksDlg::SortFunc, (DWORD_PTR)( ( it.fmt & HDF_SORTUP ) ? ( i + 1 ) : - ( i + 1 ) ) );
			m_wndList.UpdateWindow();
			break;
		}
	}
}

int CALLBACK CSageLinksDlg::SortFunc( LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort )
{
	int nRetVal = 0;

	const CLink* pData1 = (const CLink*)lParam1;
	const CLink* pData2 = (const CLink*)lParam2;

	int nColumn = abs( (int)lParamSort ) - 1;
	switch ( nColumn )
	{
	case COL_SOURCE:
		nRetVal = pData1->m_sSource.CompareNoCase( pData2->m_sSource );
		break;

	case COL_TYPE:
		nRetVal = pData1->m_nType - pData2->m_nType;
		break;

	case COL_TARGET:
		nRetVal = pData1->m_sTarget.CompareNoCase( pData2->m_sTarget );
		break;

	case COL_RESULT:
		nRetVal = pData1->m_bResult - pData2->m_bResult;
		break;
	}

	if ( nRetVal == 0 )
		nRetVal = pData1->m_bResult - pData2->m_bResult;
	if ( nRetVal == 0 )
		nRetVal = pData1->m_nType - pData2->m_nType;
	if ( nRetVal == 0 )
		nRetVal = pData1->m_sSource.CompareNoCase( pData2->m_sSource );
	if ( nRetVal == 0 )
		nRetVal = pData1->m_sTarget.CompareNoCase( pData2->m_sTarget );

	return ( lParamSort > 0 ) ? nRetVal : - nRetVal;
}

void CSageLinksDlg::ClearList()
{
	CSingleLock oLock( &m_pSection, TRUE );

	while ( ! m_pIncoming.IsEmpty() )
	{
		delete m_pIncoming.RemoveHead();
	}

	for ( int i = 0; i < m_nTotal; ++i )
	{
		delete (CLink*)m_wndList.GetItemData( i );
	}

	m_wndList.DeleteAllItems();

	m_nTotal = 0;
	m_nBad = 0;
}

BOOL CSageLinksDlg::OnKeyDown(UINT nChar, UINT /*nRepCnt*/, UINT /*nFlags*/)
{
	if ( nChar == VK_DELETE )
	{
		if ( GetFocus() == static_cast< CWnd*>( &m_wndList ) )
		{
			int nCount = 0;
			for ( int i = 0; i < m_nTotal; ++i )
			{
				if ( m_wndList.GetItemState( i, LVIS_SELECTED ) == LVIS_SELECTED )
					++ nCount;
			}

			if ( nCount )
			{
				CString sPrompt;
				sPrompt.Format( IDS_DELETE_CONFIRM, nCount );
				if ( AfxMessageBox( sPrompt, MB_YESNO | MB_ICONQUESTION ) == IDYES )
				{
					CWaitCursor wc;

					for ( int i = m_nTotal - 1; i >= 0; -- i )
					{
						if ( m_wndList.GetItemState( i, LVIS_SELECTED ) == LVIS_SELECTED )
						{
							CLink* pLink = (CLink*)m_wndList.GetItemData( i );
							if ( DeleteFile( pLink->m_sSource ) )
							{
								VERIFY( m_wndList.DeleteItem( i ) );

								-- m_nTotal;
								if ( ! pLink->m_bResult )
									-- m_nBad;

								delete pLink;
							}
						}
					}
				}
			}
			return TRUE;
		}
	}
	return FALSE;
}
