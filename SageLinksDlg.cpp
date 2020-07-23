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

#define CX_ICON		22	// pixels

// Column numbers

#define COL_SOURCE	0
#define COL_TYPE	1
#define COL_TARGET	2
#define COL_RESULT	3

// CSageLinksDlg dialog

CSageLinksDlg::CSageLinksDlg(CWnd* pParent /*=NULL*/)
	: CDialogExSized	( IDD, pParent )
	, m_sPath			( theApp.GetProfileString( OPT_SECTION, OPT_PATH, _T("C:\\") ) )
	, m_hIcon			( AfxGetApp()->LoadIcon( IDR_MAINFRAME ) )
	, m_hThread			( nullptr )
	, m_pFlag			( FALSE, TRUE )
	, m_nBad			( 0 )
	, m_nSort			( 0 )
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
	ON_NOTIFY( NM_DBLCLK, IDC_LIST, &CSageLinksDlg::OnNMDblclkList )
	ON_NOTIFY( NM_RCLICK, IDC_LIST, &CSageLinksDlg::OnNMRClickList )
	ON_NOTIFY( HDN_ITEMCLICK, 0, &CSageLinksDlg::OnHdnItemclickList )
	ON_NOTIFY( LVN_GETDISPINFO, IDC_LIST, &CSageLinksDlg::OnLvnGetdispinfoList )
	ON_NOTIFY( LVN_ODCACHEHINT, IDC_LIST, &CSageLinksDlg::OnLvnOdcachehintList )
	ON_NOTIFY( LVN_ODFINDITEM, IDC_LIST, &CSageLinksDlg::OnLvnOdfinditemList )
	ON_BN_CLICKED( IDC_DELETE, &CSageLinksDlg::OnBnClickedDelete )
END_MESSAGE_MAP()

// CSageLinksDlg message handlers

BOOL CSageLinksDlg::OnInitDialog()
{
	__super::OnInitDialog();

	SetIcon( m_hIcon, TRUE );		// Set big icon
	SetIcon( m_hIcon, FALSE );		// Set small icon

	m_oImages.Create( 16, 16, ILC_COLOR32 | ILC_MASK, 0, 100 );
	m_nImageSuccess = m_oImages.Add( (HICON)LoadImage( AfxGetResourceHandle(), MAKEINTRESOURCE( IDI_RESULT_SUCCESS ), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR | LR_SHARED ) );
	m_nImageError = m_oImages.Add( (HICON)LoadImage( AfxGetResourceHandle(), MAKEINTRESOURCE( IDI_RESULT_ERROR ), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR | LR_SHARED ) );
	m_nImageUnknown = m_oImages.Add( (HICON)LoadImage( AfxGetResourceHandle(), MAKEINTRESOURCE( IDI_UNKNOWN ), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR | LR_SHARED ) );
	m_nImageSymbolic = m_oImages.Add( (HICON)LoadImage( AfxGetResourceHandle(), MAKEINTRESOURCE( IDI_SYMBOLIC ), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR | LR_SHARED ) );
	m_nImageJunction = m_oImages.Add( (HICON)LoadImage( AfxGetResourceHandle(), MAKEINTRESOURCE( IDI_JUNCTION ), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR | LR_SHARED ) );
	m_nImageShortcut = m_oImages.Add( (HICON)LoadImage( AfxGetResourceHandle(), MAKEINTRESOURCE( IDI_SHORTCUT ), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR | LR_SHARED ) );

	ASSERT( m_wndList.GetStyle() & LVS_OWNERDATA );
	m_wndList.SetImageList( &m_oImages, LVSIL_SMALL );
	m_wndList.SetExtendedStyle( m_wndList.GetExtendedStyle() | LVS_EX_DOUBLEBUFFER | LVS_EX_FULLROWSELECT | LVS_EX_LABELTIP | LVS_EX_SUBITEMIMAGES );

	m_wndList.InsertColumn( COL_SOURCE, _T( "Source" ), LVCFMT_LEFT );
	m_wndList.InsertColumn( COL_TYPE,   _T( "" ), LVCFMT_CENTER );
	m_wndList.InsertColumn( COL_TARGET, _T( "Target" ), LVCFMT_LEFT );
	m_wndList.InsertColumn( COL_RESULT, _T( "Result" ), LVCFMT_LEFT );

	HDITEM it = { HDI_FORMAT };
	m_wndList.GetHeaderCtrl()->GetItem( COL_RESULT, &it );
	it.fmt |= HDF_SORTDOWN;
	m_wndList.GetHeaderCtrl()->SetItem( COL_RESULT, &it );
	m_nSort = - ( COL_RESULT + 1 );

	m_wndBrowse.SetWindowText( m_sPath );

	if ( ( GetKeyState( VK_SHIFT ) & 0x8000 ) == 0 )
	{
		RestoreWindowPlacement();
	}

	SetTimer( ID_TIMER, 100, nullptr );

	Start();

	Resize();

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
		m_hThread = nullptr;
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
		m_hThread = (HANDLE)_beginthreadex( nullptr, 0, &CSageLinksDlg::ThreadStub, this, 0, &id );
	}
}

unsigned __stdcall CSageLinksDlg::ThreadStub(void* param)
{
	CSageLinksDlg* pThis = (CSageLinksDlg*)param;
	pThis->Thread();
	return 0;
}

void CSageLinksDlg::OnNewItem(CLink* link)
{
	CSingleLock oLock( &m_pSection, TRUE );

	m_pList.emplace_back( link );

	if ( ! link->m_bResult )
	{
		++m_nBad;
	}
}

void CSageLinksDlg::OnTimer(UINT_PTR nIDEvent)
{
	GetDlgItem( IDC_DELETE )->EnableWindow( m_wndList.GetFirstSelectedItemPosition() != nullptr );

	CSingleLock oLock( &m_pSection, TRUE );

	const UINT nOldTotal = m_wndList.GetItemCount();
	const BOOL bIsLastVisible = ( nOldTotal == 0 ) || (BOOL)ListView_IsItemVisible( m_wndList.GetSafeHwnd(), nOldTotal - 1 );

	const UINT nTotal = m_pList.size();
	if ( nOldTotal != nTotal )
	{
		SortList();

		m_wndList.SetItemCountEx( nTotal, LVSICF_NOINVALIDATEALL | LVSICF_NOSCROLL );

		if ( bIsLastVisible )
		{
			m_wndList.EnsureVisible( nTotal - 1, TRUE );
		}
	}

	if ( m_sStatus != m_sOldStatus )
	{
		m_sOldStatus = m_sStatus;

		if ( ! m_sStatus.IsEmpty() )
		{
			CString sText;
			sText.Format( IDS_WORK, (LPCTSTR)m_sStatus );
			m_wndStatus.SetWindowText( sText );
		}
	}

	if ( nIDEvent == ID_DONE )
	{
		// Done
		m_oDirs.RemoveAll();
		m_sStatus.Empty();

		CString sOk;
		sOk.LoadString( IDS_START );
		GetDlgItem( IDOK )->SetWindowText( sOk );

		CString sText;
		sText.Format( IDS_DONE, (int)m_pList.size(), m_nBad );
		m_wndStatus.SetWindowText( sText );

		// Clean-up
		Stop();

		return;
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

void CSageLinksDlg::Resize()
{
	CRect rc;
	m_wndList.GetClientRect( &rc );
	const int width = ( rc.Width() - CX_ICON - GetSystemMetrics( SM_CXVSCROLL ) - 5 ) / 5;
	m_wndList.SetColumnWidth( COL_SOURCE, width * 2 );
	m_wndList.SetColumnWidth( COL_TYPE, CX_ICON );
	m_wndList.SetColumnWidth( COL_TARGET, width * 2 );
	m_wndList.SetColumnWidth( COL_RESULT, width );
}

void CSageLinksDlg::OnSize( UINT nType, int cx, int cy )
{
	if ( m_wndList.m_hWnd )
	{
		Resize();
	}

	__super::OnSize( nType, cx, cy );
}

void CSageLinksDlg::OnNMDblclkList(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast< LPNMITEMACTIVATE >( pNMHDR );
	*pResult = 0;

	CSingleLock oLock( &m_pSection, TRUE );

	if ( pNMItemActivate->iItem >= 0 && pNMItemActivate->iItem < (int)m_pList.size() )
	{
		const CLink* plink = m_pList.at( pNMItemActivate->iItem );

		CWaitCursor wc;

		ShellExecute( GetSafeHwnd(), nullptr, _T("explorer.exe"), _T("/select, \"") + plink->m_sSource + _T("\""), nullptr, SW_SHOWNORMAL );
	}
}

void CSageLinksDlg::OnNMRClickList(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMITEMACTIVATE pNMItemActivate = reinterpret_cast<LPNMITEMACTIVATE>( pNMHDR );
	*pResult = 0;

	CWaitCursor wc;

	CSingleLock oLock( &m_pSection, TRUE );

	// Count selected items
	CStringList oFiles;
	std::vector< int > oIndexes;
	for ( POSITION pos = m_wndList.GetFirstSelectedItemPosition(); pos; )
	{
		const int index = m_wndList.GetNextSelectedItem( pos );
		oIndexes.push_back( index );

		const CLink* plink = m_pList.at( index );
		oFiles.AddTail( plink->m_sSource );
	}

	// From end to start
	std::sort( oIndexes.begin(), oIndexes.end(), std::greater< int >() );

	if ( oFiles.GetCount() )
	{
		// Show user menu
		CMenu oMenu;
		if ( oMenu.CreatePopupMenu() )
		{
			CPoint ptAction( pNMItemActivate->ptAction );
			m_wndList.ClientToScreen( &ptAction );
			DoExplorerMenu( GetSafeHwnd(), oFiles, ptAction, oMenu.GetSafeHmenu() );

			// Re-check list for deleted items
			for ( auto i : oIndexes )
			{
				const CLink* plink = m_pList.at( i );
				if ( ! plink->IsExist() )
				{
					if ( ! plink->m_bResult )
					{
						-- m_nBad;
					}

					delete plink;

					m_pList.erase( m_pList.begin() + i );

					m_wndList.SetItemState( i, 0, LVIS_SELECTED );
				}
			}

			m_wndList.SetItemCountEx( m_pList.size(), LVSICF_NOINVALIDATEALL | LVSICF_NOSCROLL );
			m_wndList.InvalidateRect( nullptr );
		}
	}
}

void CSageLinksDlg::OnHdnItemclickList( NMHDR *pNMHDR, LRESULT *pResult )
{
	NMLISTVIEW* pLV = (NMLISTVIEW*)pNMHDR;

	m_nSort = 0;

	CHeaderCtrl* pHeader = m_wndList.GetHeaderCtrl();
	const int count = pHeader->GetItemCount();
	HDITEM it = { HDI_FORMAT };
	for ( int i = 0; i < count; ++i )
	{
		pHeader->GetItem( i, &it );
		if ( i == pLV->iItem )
		{
			it.fmt = ( it.fmt & ~( HDF_SORTDOWN | HDF_SORTUP ) ) | ( ( it.fmt & HDF_SORTUP ) ? HDF_SORTDOWN : HDF_SORTUP );

			m_nSort = ( it.fmt & HDF_SORTDOWN ) ? ( i + 1 ) : - ( i + 1 );
		}
		else
		{
			it.fmt &= ~( HDF_SORTDOWN | HDF_SORTUP );
		}
		pHeader->SetItem( i, &it );
	}

	SortList();

	*pResult = 0;
}

void CSageLinksDlg::SortList()
{
	if ( m_nSort )
	{
		std::sort( m_pList.begin(), m_pList.end(),
			[=](const CLink* pData1, const CLink* pData2) noexcept
			{
				int nRetVal = 0;

				const int nColumn = abs( m_nSort ) - 1;
				switch ( nColumn )
				{
				case COL_SOURCE:
					nRetVal = pData1->m_sSource.CompareNoCase( pData2->m_sSource );
					break;

				case COL_TYPE:
					nRetVal = (int)pData1->m_nType - (int)pData2->m_nType;
					break;

				case COL_TARGET:
					nRetVal = pData1->m_sTarget.CompareNoCase( pData2->m_sTarget );
					break;

				case COL_RESULT:
					nRetVal = pData1->m_sResult.CompareNoCase( pData2->m_sResult );
					break;
				}

				if ( nRetVal == 0 )
					nRetVal = pData1->m_sResult.CompareNoCase( pData2->m_sResult );
				if ( nRetVal == 0 )
					nRetVal = (int)pData1->m_nType - (int)pData2->m_nType;
				if ( nRetVal == 0 )
					nRetVal = pData1->m_sSource.CompareNoCase( pData2->m_sSource );
				if ( nRetVal == 0 )
					nRetVal = pData1->m_sTarget.CompareNoCase( pData2->m_sTarget );

				return ( ( ( m_nSort > 0 ) ? nRetVal : ( - nRetVal ) ) > 0 );
			} );

		m_wndList.InvalidateRect( nullptr );
	}
}

void CSageLinksDlg::ClearList()
{
	CSingleLock oLock( &m_pSection, TRUE );

	for ( auto i : m_pList )
	{
		delete i;
	}

	m_pList.clear();

	m_wndList.DeleteAllItems();

	m_nBad = 0;
}

void CSageLinksDlg::OnBnClickedDelete()
{
	CSingleLock oLock( &m_pSection, TRUE );

	// Count selected items
	std::vector< int > oIndexes;
	for ( POSITION pos = m_wndList.GetFirstSelectedItemPosition(); pos; )
	{
		oIndexes.push_back( m_wndList.GetNextSelectedItem( pos ) );
	}

	// From end to start
	std::sort( oIndexes.begin(), oIndexes.end(), std::greater< int >() );

	// Ask user
	if ( UINT nCount = oIndexes.size() )
	{
		CString sPrompt;
		sPrompt.Format( IDS_DELETE_CONFIRM, nCount );
		if ( AfxMessageBox( sPrompt, MB_YESNO | MB_ICONQUESTION ) == IDYES )
		{
			// Delete selected items

			CWaitCursor wc;

			for ( auto i : oIndexes )
			{
				const CLink* plink = m_pList.at( i );
				if ( plink->DeleteFile() )
				{
					if ( ! plink->m_bResult )
					{
						--m_nBad;
					}

					delete plink;

					m_pList.erase( m_pList.begin() + i );

					m_wndList.SetItemState( i, 0, LVIS_SELECTED );
				}
			}

			m_wndList.SetItemCountEx( m_pList.size(), LVSICF_NOINVALIDATEALL | LVSICF_NOSCROLL );
			m_wndList.InvalidateRect( nullptr );
		}
	}
}

BOOL CSageLinksDlg::OnKeyDown(UINT nChar, UINT /*nRepCnt*/, UINT /*nFlags*/)
{
	if ( nChar == VK_F5 )
	{
		Start();

		return TRUE;
	}
	else if ( nChar == VK_DELETE )
	{
		if ( GetFocus() == static_cast< CWnd*>( &m_wndList ) )
		{
			OnBnClickedDelete();

			return TRUE;
		}
	}
	return FALSE;
}

void CSageLinksDlg::OnLvnGetdispinfoList(NMHDR *pNMHDR, LRESULT *pResult)
{
	NMLVDISPINFO* pDispInfo = reinterpret_cast< NMLVDISPINFO* >( pNMHDR );
	*pResult = 0;

	CSingleLock oLock( &m_pSection, TRUE );

	LVITEMW& item = pDispInfo->item;
	if ( item.iItem >= 0 && item.iItem < (int)m_pList.size() )
	{
		const CLink* plink = m_pList.at( item.iItem );

		if ( item.mask & LVIF_PARAM )
		{
			item.lParam = (DWORD_PTR)item.iItem;
		}

		switch ( item.iSubItem )
		{
		case COL_SOURCE:
			if ( item.mask & LVIF_IMAGE )
			{
				item.iImage = ( plink->m_hIcon ? m_oImages.Add( plink->m_hIcon ) : m_nImageUnknown );
			}
			if ( item.mask & LVIF_TEXT )
			{
				_tcsncpy_s( item.pszText, item.cchTextMax, plink->m_sSource, _TRUNCATE );
			}
			break;

		case COL_TYPE:
			if ( item.mask & LVIF_IMAGE )
			{
				item.iImage =
					( ( plink->m_nType == LinkType::Symbolic ) ? m_nImageSymbolic :
					( ( plink->m_nType == LinkType::Junction ) ? m_nImageJunction :
					( ( plink->m_nType == LinkType::Shortcut ) ? m_nImageShortcut : m_nImageUnknown ) ) );
			}
			break;

		case COL_TARGET:
			if ( item.mask & LVIF_IMAGE )
			{
				item.iImage = plink->m_bResult ? m_nImageSuccess : m_nImageError;
			}
			if ( item.mask & LVIF_TEXT )
			{
				_tcsncpy_s( item.pszText, item.cchTextMax, plink->m_sTarget, _TRUNCATE );
			}
			break;

		case COL_RESULT:
			if ( item.mask & LVIF_TEXT )
			{
				_tcsncpy_s( item.pszText, item.cchTextMax, plink->m_sResult, _TRUNCATE );
			}
			break;
		}
	}

	item.mask |= LVIF_DI_SETITEM;
}

void CSageLinksDlg::OnLvnOdcachehintList(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVCACHEHINT pCacheHint = reinterpret_cast<LPNMLVCACHEHINT>( pNMHDR );
	*pResult = 0;
	pCacheHint;
}

void CSageLinksDlg::OnLvnOdfinditemList(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLVFINDITEM pFindInfo = reinterpret_cast<LPNMLVFINDITEM>( pNMHDR );
	*pResult = 0;
	pFindInfo;
}
