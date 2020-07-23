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


#define CX_ICON		20	// pixels
#define LONG_PATH	512	// symbols

// Column numbers

#define COL_SOURCE	0
#define COL_TYPE	1
#define COL_TARGET	2
#define COL_RESULT	3


class ATL_NO_VTABLE CShellItem
{
public:
	CShellItem(LPCTSTR szFullPath, void** ppFolder = NULL, HWND hWnd = NULL)
		: m_pidl	( NULL )
		, m_pLastId	( NULL )
	{
		CComPtr< IShellFolder > pDesktop;
		HRESULT hr = SHGetDesktopFolder( &pDesktop );
		if ( FAILED( hr ) )
			return;

		hr = pDesktop->ParseDisplayName( hWnd, 0, CT2OLE( szFullPath ), NULL, &m_pidl, NULL );
		if ( FAILED( hr ) )
			return;

		m_pLastId = ILFindLastID( m_pidl );

		if ( ppFolder )
		{
			USHORT temp = m_pLastId->mkid.cb;
			m_pLastId->mkid.cb = 0;
			hr = pDesktop->BindToObject( m_pidl, NULL, IID_IShellFolder, ppFolder );
			m_pLastId->mkid.cb = temp;
		}
	}

	~CShellItem()
	{
		if ( m_pidl )
			CoTaskMemFree( m_pidl );
	}

	inline operator LPITEMIDLIST() const
	{
		return m_pidl;
	}

public:
	PIDLIST_RELATIVE	m_pidl;		// Full path
	PUITEMID_CHILD		m_pLastId;	// Filename only
};

class CShellList : public CList< CShellItem* >
{
public:
	CShellList(const CStringList& oFiles)
		: m_pID	( NULL )
	{
		for ( POSITION pos = oFiles.GetHeadPosition(); pos; )
		{
			const CString strPath = oFiles.GetNext( pos );

			CShellItem* pItemIDList = new CShellItem( strPath, ( m_pFolder ? NULL : (void**)&m_pFolder ) );	// Get only one
			if ( ! pItemIDList )
				// Out of memory
				return;

			if ( pItemIDList->m_pidl )
				AddTail( pItemIDList );
			else
				// Bad path
				delete pItemIDList;
		}

		if ( GetCount() == 0 )
			// No files
			return;

		m_pID.Attach( new PCUITEMID_CHILD[ GetCount() ] );
		if ( ! m_pID )
			// Out of memory
			return;

		size_t i = 0;
		for ( POSITION pos = GetHeadPosition(); pos; ++ i )
			m_pID[ i ] = GetNext( pos )->m_pLastId;
	}

	virtual ~CShellList()
	{
		for ( POSITION pos = GetHeadPosition(); pos; )
			delete GetNext( pos );
		RemoveAll();
	}

	// Creates menu from file paths list
	inline bool GetMenu( HWND hWnd, void** ppContextMenu )
	{
		return m_pFolder && SUCCEEDED( m_pFolder->GetUIObjectOf( hWnd, (UINT)GetCount(), m_pID, IID_IContextMenu, NULL, ppContextMenu ) );
	}

protected:
	CComPtr< IShellFolder >				m_pFolder;	// First file folder
	CAutoVectorPtr< PCUITEMID_CHILD >	m_pID;		// File ItemID array
};

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
		LPTSTR lpszTemp = NULL;
		FormatMessage( FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM, NULL, hr, 0, (LPTSTR)&lpszTemp, 0, NULL );
		for ( int i = 0; ! lpszTemp && i < _countof( szModules ); ++i )
		{
			if ( HMODULE hModule = LoadLibraryEx( szModules[ i ], NULL, LOAD_LIBRARY_AS_DATAFILE ) )
			{
				FormatMessage( FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_HMODULE, hModule, hr, 0, (LPTSTR)&lpszTemp, 0, NULL );
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

	sError.AppendFormat( _T(" (0x%08x)"), hr );
	return sError;
}

#define ID_SHELL_MENU_MIN 40000
#define ID_SHELL_MENU_MAX 41000

inline UINT_PTR DoExplorerMenu(HWND hwnd, const CStringList& oFiles, POINT point, HMENU hMenu)
{
	UINT_PTR nCmd = 0;
	CComPtr< IContextMenu > pContextMenu1;
	CShellList oItemIDListList( oFiles );
	if ( oItemIDListList.GetMenu( hwnd, (void**)&pContextMenu1 ) )
	{
		HRESULT hr = pContextMenu1->QueryContextMenu( hMenu, 0, ID_SHELL_MENU_MIN, ID_SHELL_MENU_MAX, CMF_NORMAL | CMF_EXPLORE );
		if ( SUCCEEDED( hr ) )
		{
			::SetForegroundWindow( hwnd );
			nCmd = ::TrackPopupMenu( hMenu, TPM_RETURNCMD | TPM_LEFTALIGN | TPM_TOPALIGN |
				TPM_LEFTBUTTON | TPM_RIGHTBUTTON, point.x, point.y, 0, hwnd, NULL );
			::PostMessage( hwnd, WM_NULL, 0, 0 );

			// If a command was selected from the shell menu, execute it.
			if ( nCmd >= ID_SHELL_MENU_MIN && nCmd <= ID_SHELL_MENU_MAX )
			{
				CMINVOKECOMMANDINFOEX ici = { sizeof( CMINVOKECOMMANDINFOEX ) };
				ici.hwnd = hwnd;
				ici.lpVerb = reinterpret_cast<LPCSTR>( nCmd - ID_SHELL_MENU_MIN );
				ici.lpVerbW = reinterpret_cast<LPCWSTR>( nCmd - ID_SHELL_MENU_MIN );
				ici.nShow = SW_SHOWNORMAL;
				hr = pContextMenu1->InvokeCommand( (CMINVOKECOMMANDINFO*)&ici );
			}
		}
	}

	return nCmd;
}

// CSageLinksDlg dialog

CSageLinksDlg::CSageLinksDlg(CWnd* pParent /*=NULL*/)
	: CDialog	( IDD_SAGELINKS_DIALOG, pParent )
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

BEGIN_MESSAGE_MAP(CSageLinksDlg, CDialog)
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

	m_wndBrowse.SetWindowText( theApp.GetProfileString( OPT_SECTION, OPT_PATH, _T("C:\\") ) );

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

inline BOOL IsDots(LPCTSTR szFileName)
{
	return szFileName[ 0 ] == _T('.') && ( szFileName[ 1 ] == 0 || ( szFileName[ 1 ] == _T('.') && szFileName[ 2 ] == 0 ) );
}

inline BOOL IsLocal(LPCTSTR szFileName)
{
	return _istalpha( szFileName[ 0 ] ) && szFileName[ 1 ] == _T( ':' ) && ( szFileName[ 2 ] == 0 || szFileName[ 2 ] == _T('\\') );
}

inline BOOL IsUNC(LPCTSTR szFileName)
{
	return szFileName[ 0 ] == _T('\\') && szFileName[ 1 ] == _T('\\') && _istalnum( szFileName[ 2 ] );
}

CString GetRemotedPath(const CString& sBasePath, const CString& sLocalPath)
{
	// Source path is UNC
	int n = sBasePath.Mid( 2 ).Find( _T( '\\' ) ) + 2;
	if ( n > 2 && IsLocal( sLocalPath ) )
	{
		// Target path is local and source path is rooted UNC
		return sBasePath.Left( n + 1 ) + sLocalPath.GetAt( 0 ) + _T( '$' ) + sLocalPath.Mid( 2 );
	}
	return CString();
}

void CSageLinksDlg::Thread()
{
	HRESULT hres;
	DWORD nRead;
	CAutoVectorPtr< BYTE > pBuf( new BYTE[ LONG_PATH ] );

	hres = CoInitializeEx( NULL, COINIT_APARTMENTTHREADED );

#ifndef _WIN64
	typedef BOOL (WINAPI *tWow64DisableWow64FsRedirection)(PVOID*);
	typedef BOOL (WINAPI *tWow64RevertWow64FsRedirection)(PVOID);
	HMODULE hKernel32 = LoadLibrary( _T("Kernel32.dll") );
	tWow64DisableWow64FsRedirection fnWow64DisableWow64FsRedirection =
		(tWow64DisableWow64FsRedirection)GetProcAddress( hKernel32, "Wow64DisableWow64FsRedirection" );
	tWow64RevertWow64FsRedirection fnWow64RevertWow64FsRedirection =
		(tWow64RevertWow64FsRedirection)GetProcAddress( hKernel32, "Wow64RevertWow64FsRedirection" );
	PVOID pRedir = NULL;
	if ( fnWow64DisableWow64FsRedirection ) fnWow64DisableWow64FsRedirection( &pRedir );
#endif

	typedef HRESULT (STDAPICALLTYPE *tSHGetNameFromIDList)(PCIDLIST_ABSOLUTE, SIGDN, PWSTR*);
	HMODULE hShell32 = LoadLibrary( _T("Shell32.dll") );
	tSHGetNameFromIDList fnSHGetNameFromIDList =
		(tSHGetNameFromIDList)GetProcAddress( hShell32, "SHGetNameFromIDList" );

	CString sProgramFiles32;
	ExpandEnvironmentStrings( _T( "%ProgramFiles(x86)%" ), sProgramFiles32.GetBuffer( LONG_PATH ), LONG_PATH );
	sProgramFiles32.ReleaseBuffer();

	CString sProgramFiles64;
	ExpandEnvironmentStrings( _T( "%ProgramW6432%" ), sProgramFiles64.GetBuffer( LONG_PATH ), LONG_PATH );
	sProgramFiles64.ReleaseBuffer();

	if ( sProgramFiles64.GetAt( 0 ) == _T( '%' ) || sProgramFiles32.CompareNoCase( sProgramFiles64 ) == 0 )
		sProgramFiles64.Empty();

	while ( ! m_oDirs.IsEmpty() && WaitForSingleObject( m_pFlag, 0 ) == WAIT_TIMEOUT )
	{
		CString sDir = m_oDirs.RemoveHead();
		sDir.TrimRight( _T("\\") );

		{
			CSingleLock oLock( &m_pSection, TRUE );

			m_sStatus = sDir;
		}

		const BOOL bUNC = IsUNC( sDir );

		WIN32_FIND_DATA wfa = {};
		HANDLE hFF = FindFirstFile( sDir + _T("\\*.*"), &wfa );
		if ( hFF != INVALID_HANDLE_VALUE )
		{
			do
			{
				if ( IsDots( wfa.cFileName ) )
					// Skip dots
					continue;

				if ( ( wfa.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT ) != 0 )
				{
					LinkType nType = LinkType::Unknown;
					switch ( wfa.dwReserved0 )
					{
					case IO_REPARSE_TAG_MOUNT_POINT:
						nType = LinkType::Junction;
						break;
					case IO_REPARSE_TAG_SYMLINK:
						nType = LinkType::Symbolic;
						break;
					}

					// Reparse point
					const CString sPath( sDir + _T( "\\" ) + wfa.cFileName );
					CString sTarget;

					BOOL bResult = FALSE;
					HANDLE hFile = CreateFile( sPath, FILE_READ_ATTRIBUTES,
						FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, NULL, OPEN_EXISTING,
						FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS, NULL );
					if ( hFile != INVALID_HANDLE_VALUE )
					{
						nRead = 0;
						if ( DeviceIoControl( hFile, FSCTL_GET_REPARSE_POINT, NULL, 0, pBuf, LONG_PATH, &nRead, NULL ) )
						{
							const REPARSE_DATA_BUFFER* pReparse = (const REPARSE_DATA_BUFFER*)(const BYTE*)pBuf;
							CString sReparse;
							if ( wfa.dwReserved0 == IO_REPARSE_TAG_MOUNT_POINT )
							{
								sReparse.Append( pReparse->MountPointReparseBuffer.PathBuffer + pReparse->MountPointReparseBuffer.SubstituteNameOffset / sizeof( WCHAR ),
									pReparse->MountPointReparseBuffer.SubstituteNameLength / sizeof( WCHAR ) );
							}
							else if ( wfa.dwReserved0 == IO_REPARSE_TAG_SYMLINK )
							{
								sReparse.Append( pReparse->SymbolicLinkReparseBuffer.PathBuffer + pReparse->SymbolicLinkReparseBuffer.SubstituteNameOffset / sizeof( WCHAR ),
									pReparse->SymbolicLinkReparseBuffer.SubstituteNameLength / sizeof( WCHAR ) );
							}

							if ( _tcsncmp( sReparse, _T("\\??\\"), 4 ) == 0 )
								sReparse = sReparse.Mid( 4 );

							if ( _tcsncmp( sReparse, _T( "\\\\?\\" ), 4 ) == 0 )
								sReparse = sReparse.Mid( 4 );

							if ( ! sReparse.IsEmpty() )
							{
								if ( ! IsLocal( sReparse ) && ! IsUNC( sReparse ) )
									// Relative path
									sReparse = sDir + _T( "\\" ) + sReparse;

								// Make pretty path
								PathCanonicalize( sTarget.GetBuffer( LONG_PATH ), sReparse );
								sTarget.ReleaseBuffer();

								if ( bUNC )
								{
									// Source path is UNC
									const CString sRemoteTarget = GetRemotedPath( sPath, sTarget );
									if ( ! sRemoteTarget.IsEmpty() )
									{
										// Target path is local and source path is rooted UNC
										bResult = ( GetFileAttributes( sRemoteTarget ) != INVALID_FILE_ATTRIBUTES );
										if ( bResult )
											sTarget = sRemoteTarget;
									}
								}

								if ( ! bResult )
								{
									bResult = ( GetFileAttributes( sTarget ) != INVALID_FILE_ATTRIBUTES );
								}

								if ( ! bResult )
								{
									TRACE( "%s : Target error %d\n", (LPCSTR)CT2A( (LPCTSTR)sPath ), GetLastError() );
								}
							}
							else
							{
								TRACE( "%s : No Target\n", (LPCSTR)CT2A( (LPCTSTR)sPath ) );
							}
						}
						else
						{
							sTarget = ErrorMessage( GetLastError() );
							TRACE( "%s : IO Error %d\n", (LPCSTR)CT2A( (LPCTSTR)sPath ), GetLastError() );
						}

						CloseHandle( hFile );
					}
					else
					{
						sTarget = ErrorMessage( GetLastError() );
						TRACE( "%s : Open Error %d\n", (LPCSTR)CT2A( (LPCTSTR)sPath ), GetLastError() );
					}

					if ( nType != LinkType::Unknown )
					{
						SHFILEINFO sfi = {};
						SHGetFileInfo( bResult ? sTarget : sPath, 0, &sfi, sizeof( sfi ), SHGFI_ICON | SHGFI_SMALLICON );

						CSingleLock oLock( &m_pSection, TRUE );

						m_pIncoming.AddTail( new CLink( nType, sfi.hIcon, sPath, sTarget, bResult ) );
					}
				}
				else if ( ( wfa.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY ) != 0 )
				{
					// Directory
					const CString sPath( sDir + _T( "\\" ) + wfa.cFileName );
					m_oDirs.AddTail( sPath );
				}
				else
				{
					const size_t len = _tcslen( wfa.cFileName );
					if ( len > 4 && _tcscmp( wfa.cFileName + len - 4, _T( ".lnk" ) ) == 0 )
					{
						// Shortcut
						const CString sPath( sDir + _T( "\\" ) + wfa.cFileName );
						CString sTarget;
						BOOL bResult = FALSE;

						LinkType nType = LinkType::Shortcut;

						CComPtr< IShellLink > pShellLink;
						if ( SUCCEEDED( hres = pShellLink.CoCreateInstance( CLSID_ShellLink ) ) )
						{
							CComQIPtr< IPersistFile > pLinkFile( pShellLink );
							if ( pLinkFile && SUCCEEDED( hres = pLinkFile->Load( sPath, STGM_READ ) ) )
							{
								CString sLinkPath;
								hres = pShellLink->GetPath( sLinkPath.GetBuffer( LONG_PATH ), LONG_PATH, NULL, SLGP_RAWPATH );
								sLinkPath.ReleaseBuffer();
								if ( hres == S_OK )
								{
									// Expand environment if any
									CString sExpanded;
									ExpandEnvironmentStrings( sLinkPath, sExpanded.GetBuffer( LONG_PATH ), LONG_PATH );
									sExpanded.ReleaseBuffer();

									// Make pretty path
									PathCanonicalize( sTarget.GetBuffer( LONG_PATH ), sExpanded );
									sTarget.ReleaseBuffer();

									if ( bUNC )
									{
										// Source path is UNC
										const CString sRemoteTarget = GetRemotedPath( sPath, sTarget );
										if ( ! sRemoteTarget.IsEmpty() )
										{
											// Target path is local and source path is rooted UNC
											bResult = ( GetFileAttributes( sRemoteTarget ) != INVALID_FILE_ATTRIBUTES );
											if ( bResult )
												sTarget = sRemoteTarget;

											// Try changing 32-bit to 64-bit Program Files
											if ( ! bResult && ! sProgramFiles64.IsEmpty() && _tcsncicmp( sTarget, sProgramFiles32, sProgramFiles32.GetLength() ) == 0 )
											{
												const CString sTarget64 = sProgramFiles64 + sTarget.Mid( sProgramFiles32.GetLength() );
												const CString sRemoteTarget64 = GetRemotedPath( sPath, sTarget64 );
												bResult = ( GetFileAttributes( sRemoteTarget64 ) != INVALID_FILE_ATTRIBUTES );
												if ( bResult )
													sTarget = sRemoteTarget64;
											}
										}
										// else Target path is UNC too
									}
									// else Source path is local

									if ( ! bResult )
									{
										bResult = ( GetFileAttributes( sTarget ) != INVALID_FILE_ATTRIBUTES );

										// Try changing 32-bit to 64-bit Program Files
										if ( ! bResult && ! sProgramFiles64.IsEmpty() && _tcsncicmp( sTarget, sProgramFiles32, sProgramFiles32.GetLength() ) == 0 )
										{
											const CString sTarget64 = sProgramFiles64 + sTarget.Mid( sProgramFiles32.GetLength() );
											bResult = ( GetFileAttributes( sTarget64 ) != INVALID_FILE_ATTRIBUTES );
											if ( bResult )
												sTarget = sTarget64;
										}
									}

									if ( ! bResult )
									{
										TRACE( "%s : Target result %d\n", (LPCSTR)CT2A( (LPCTSTR)sPath ), GetLastError() );

										// Show raw path
										sTarget = sLinkPath;
									}
								}
								else if ( hres == S_FALSE )
								{
									// Non-file link
									PIDLIST_ABSOLUTE pidl = NULL;
									hres = pShellLink->GetIDList( &pidl );
									if ( hres == S_OK )
									{
										LPWSTR pName = NULL;
										if ( fnSHGetNameFromIDList )
										{
											hres = fnSHGetNameFromIDList( pidl, SIGDN_NORMALDISPLAY, &pName );
											if ( hres == S_OK )
											{
												sTarget = _T( "\"" ) + CString( pName ) + _T( "\"" );
												CoTaskMemFree( pName );
											}
										}
										CoTaskMemFree( pidl );

										bResult = TRUE;
										TRACE( "%s : Non-file shell link\n", (LPCSTR)CT2A( (LPCTSTR)sPath ) );
									}
									else if ( hres == S_FALSE )
									{
										// Unknown shortcut type
										nType = LinkType::Unknown;
									}
									else
									{
										sTarget = ErrorMessage( hres );
										TRACE( "%s : Failed GetIDList()\n", (LPCSTR)CT2A( (LPCTSTR)sPath ) );
									}
								}
								else
								{
									sTarget = ErrorMessage( hres );
									TRACE( "%s : Failed GetPath()\n", (LPCSTR)CT2A( (LPCTSTR)sPath ) );
								}
							}
							else
							{
								sTarget = ErrorMessage( hres );
								TRACE( "%s : Bad shell link\n", (LPCSTR)CT2A( (LPCTSTR)sPath ) );
							}
						}
						else
						{
							sTarget = ErrorMessage( hres );
							TRACE( "%s : Failed CoCreateInstance()\n", (LPCSTR)CT2A( (LPCTSTR)sPath ) );
						}

						if ( nType != LinkType::Unknown )
						{
							SHFILEINFO sfi = {};
							SHGetFileInfo( sPath, 0, &sfi, sizeof( sfi ), SHGFI_ICON | SHGFI_SMALLICON );

							CSingleLock oLock( &m_pSection, TRUE );

							m_pIncoming.AddTail( new CLink( nType, sfi.hIcon, sPath, sTarget, bResult ) );
						}
					}
				}
			}
			while( FindNextFile( hFF, &wfa ) && WaitForSingleObject( m_pFlag, 0 ) == WAIT_TIMEOUT );

			FindClose( hFF );
		}
	}

	CoUninitialize();

#ifndef _WIN64
	if ( fnWow64RevertWow64FsRedirection ) fnWow64RevertWow64FsRedirection( pRedir );
#endif

	CSingleLock oLock( &m_pSection, TRUE );

	m_oDirs.RemoveAll();
	m_sStatus.Empty();

	m_pIncoming.AddTail( new CLink() );
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
