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

void CSageLinksDlg::Thread()
{
	HRESULT hres;
	DWORD nRead;
	CAutoVectorPtr< BYTE > pBuf( new BYTE[ 4096 ] );

	hres = CoInitializeEx( nullptr, COINIT_APARTMENTTHREADED );

#ifndef _WIN64
	typedef BOOL (WINAPI *tWow64DisableWow64FsRedirection)(PVOID*);
	typedef BOOL (WINAPI *tWow64RevertWow64FsRedirection)(PVOID);
	HMODULE hKernel32 = LoadLibrary( _T("Kernel32.dll") );
	tWow64DisableWow64FsRedirection fnWow64DisableWow64FsRedirection =
		(tWow64DisableWow64FsRedirection)GetProcAddress( hKernel32, "Wow64DisableWow64FsRedirection" );
	tWow64RevertWow64FsRedirection fnWow64RevertWow64FsRedirection =
		(tWow64RevertWow64FsRedirection)GetProcAddress( hKernel32, "Wow64RevertWow64FsRedirection" );
	PVOID pRedir = nullptr;
	if ( fnWow64DisableWow64FsRedirection ) fnWow64DisableWow64FsRedirection( &pRedir );
#endif

	typedef HRESULT (STDAPICALLTYPE *tSHGetNameFromIDList)(PCIDLIST_ABSOLUTE, SIGDN, PWSTR*);
	HMODULE hShell32 = LoadLibrary( _T("Shell32.dll") );
	tSHGetNameFromIDList fnSHGetNameFromIDList =
		(tSHGetNameFromIDList)GetProcAddress( hShell32, "SHGetNameFromIDList" );

	CString sProgramFiles32;
	ExpandEnvironmentStrings( _T( "%ProgramFiles(x86)%" ), sProgramFiles32.GetBuffer( MAX_PATH ), MAX_PATH );
	sProgramFiles32.ReleaseBuffer();

	CString sProgramFiles64;
	ExpandEnvironmentStrings( _T( "%ProgramW6432%" ), sProgramFiles64.GetBuffer( MAX_PATH ), MAX_PATH );
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
		HANDLE hFF = FindFirstFile( ( bUNC ? sDir : ( LONG_PREFIX + sDir ) ) + _T("\\*.*"), &wfa );
		if ( hFF != INVALID_HANDLE_VALUE )
		{
			do
			{
				if ( IsDots( wfa.cFileName ) )
					// Skip dots
					continue;

				const CString sPath( sDir + _T( "\\" ) + wfa.cFileName );
				CString sTarget, sResult;
				BOOL bResult = FALSE;

				if ( ( wfa.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT ) != 0 )
				{
					// Reparse point
					LinkType nType;
					switch ( wfa.dwReserved0 )
					{
					case IO_REPARSE_TAG_MOUNT_POINT:
						nType = LinkType::Junction;
						break;

					case IO_REPARSE_TAG_SYMLINK:
						nType = LinkType::Symbolic;
						break;

					default:
						nType = LinkType::Unknown;
					}

					HANDLE hFile = CreateFile( ( bUNC ? sPath : ( LONG_PREFIX + sPath ) ), FILE_READ_ATTRIBUTES,
						FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING,
						FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS, nullptr );
					if ( hFile != INVALID_HANDLE_VALUE )
					{
						nRead = 0;
						if ( DeviceIoControl( hFile, FSCTL_GET_REPARSE_POINT, nullptr, 0, pBuf, 4096, &nRead, nullptr ) )
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
							else
							{
								sResult.Format( IDS_UNKNOWN_REPARSE, wfa.dwReserved0 );
							}

							if ( _tcsncmp( sReparse, _P( UNC_PREFIX ) ) == 0 )
							{
								sReparse = _T("\\\\") + sReparse.Mid( 8 );
							}
							else if ( _tcsncmp( sReparse, _P( TAG_PREFIX ) ) == 0 )
							{
								sReparse = sReparse.Mid( 4 );
							}

							if ( _tcsncmp( sReparse, _P( LONG_PREFIX ) ) == 0 )
							{
								sReparse = sReparse.Mid( 4 );
							}

							if ( ! sReparse.IsEmpty() )
							{
								if ( ! IsLocal( sReparse ) && ! IsUNC( sReparse ) )
								{
									// Relative path
									sReparse = sDir + _T( "\\" ) + sReparse;
								}

								sTarget = sReparse;

								if ( bUNC )
								{
									// Source path is UNC
									const CString sRemoteTarget = GetRemotedPath( sPath, sTarget );
									if ( ! sRemoteTarget.IsEmpty() )
									{
										// Target path is local and source path is rooted UNC
										bResult = IsExist( sRemoteTarget );
										if ( bResult )
										{
											sTarget = sRemoteTarget;
										}
									}
								}

								if ( ! bResult )
								{
									bResult = IsExist( sTarget );
									if ( ! bResult  )
									{
										bResult = IsExist( LONG_PREFIX + sTarget );
									}

								}

								if ( ! bResult )
								{
									sResult = ErrorMessage( GetLastError() );
									TRACE( _T("Target error %s : %s\n"), (LPCTSTR)sResult, (LPCTSTR)sPath );
								}
							}
							else
							{
								TRACE( _T("No Target : %s\n"), (LPCTSTR)sPath );
							}
						}
						else
						{
							sResult = ErrorMessage( GetLastError() );
							TRACE( _T("IO Error %s : %s\n"), (LPCTSTR)sResult, (LPCTSTR)sPath );
						}

						CloseHandle( hFile );
					}
					else
					{
						sResult = ErrorMessage( GetLastError() );
						TRACE( _T("Open Error %s : %s\n"), (LPCTSTR)sResult, (LPCTSTR)sPath );
					}

					if ( nType != LinkType::Unknown )
					{
						SHFILEINFO sfi = {};
						SHGetFileInfo( bResult ? sTarget : sPath, 0, &sfi, sizeof( sfi ), SHGFI_ICON | SHGFI_SMALLICON );
						OnNewItem( new CLink( nType, sfi.hIcon, sPath, sTarget, sResult, bResult ) );
					}
				}
				else if ( ( wfa.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY ) != 0 )
				{
					// Directory
					m_oDirs.AddTail( sPath );
				}
				else
				{
					const size_t len = _tcsnlen( wfa.cFileName, _countof( wfa.cFileName ) );
					if ( len > 4 && _tcscmp( wfa.cFileName + len - 4, _T( ".lnk" ) ) == 0 )
					{
						// Shortcut
						LinkType nType = LinkType::Shortcut;

						CComPtr< IShellLink > pShellLink;
						if ( SUCCEEDED( hres = pShellLink.CoCreateInstance( CLSID_ShellLink ) ) )
						{
							CComQIPtr< IPersistFile > pLinkFile( pShellLink );
							if ( pLinkFile &&
								( SUCCEEDED( hres = pLinkFile->Load( sPath, STGM_READ ) ) ||
								  SUCCEEDED( hres = pLinkFile->Load( LONG_PREFIX + sPath, STGM_READ ) ) ) )
							{
								CString sLinkPath;
								hres = pShellLink->GetPath( sLinkPath.GetBuffer( LONG_PATH ), LONG_PATH, nullptr, SLGP_RAWPATH );
								sLinkPath.ReleaseBuffer();
								if ( hres == S_OK )
								{
									// Expand environment if any
									CString sExpanded;
									ExpandEnvironmentStrings( sLinkPath, sExpanded.GetBuffer( LONG_PATH ), LONG_PATH );
									sExpanded.ReleaseBuffer();

									sTarget = sExpanded;

									if ( bUNC )
									{
										// Source path is UNC
										const CString sRemoteTarget = GetRemotedPath( sPath, sTarget );
										if ( ! sRemoteTarget.IsEmpty() )
										{
											// Target path is local and source path is rooted UNC
											bResult = IsExist( sRemoteTarget );
											if ( bResult )
											{
												sTarget = sRemoteTarget;
											}

											// Try changing 32-bit to 64-bit Program Files
											if ( ! bResult && ! sProgramFiles64.IsEmpty() && _tcsncicmp( sTarget, sProgramFiles32, sProgramFiles32.GetLength() ) == 0 )
											{
												const CString sTarget64 = sProgramFiles64 + sTarget.Mid( sProgramFiles32.GetLength() );
												const CString sRemoteTarget64 = GetRemotedPath( sPath, sTarget64 );
												bResult = IsExist( sRemoteTarget64 );
												if ( bResult )
												{
													sTarget = sRemoteTarget64;
												}
											}
										}
										// else Target path is UNC too
									}
									// else Source path is local

									if ( ! bResult )
									{
										bResult = IsExist( sTarget );
										if ( ! bResult )
										{
											bResult = IsExist( LONG_PREFIX + sTarget );
										}

										// Try changing 32-bit to 64-bit Program Files
										if ( ! bResult && ! sProgramFiles64.IsEmpty() && _tcsncicmp( sTarget, sProgramFiles32, sProgramFiles32.GetLength() ) == 0 )
										{
											const CString sTarget64 = sProgramFiles64 + sTarget.Mid( sProgramFiles32.GetLength() );
											bResult = IsExist( sTarget64 );
											if ( bResult )
											{
												sTarget = sTarget64;
											}
											else
											{
												bResult = IsExist( LONG_PREFIX + sTarget64 );
												if ( bResult )
												{
													sTarget = sTarget64;
												}
											}
										}
									}

									if ( ! bResult )
									{
										sResult = ErrorMessage( GetLastError() );
										TRACE( _T("Target error %s : %s\n"), (LPCTSTR)sResult, (LPCTSTR)sPath );

										// Show raw path
										sTarget = sLinkPath;
									}
								}
								else if ( hres == S_FALSE )
								{
									// Non-file link
									PIDLIST_ABSOLUTE pidl = nullptr;
									hres = pShellLink->GetIDList( &pidl );
									if ( hres == S_OK )
									{
										LPWSTR pName = nullptr;
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
										TRACE( _T("Non-file shell link : %s\n"), (LPCTSTR)sPath );
									}
									else if ( hres == S_FALSE )
									{
										// Unknown shortcut type
										nType = LinkType::Unknown;
									}
									else
									{
										sResult = ErrorMessage( hres );
										TRACE( _T("GetIDList() error %s : %s\n"), (LPCTSTR)sResult, (LPCTSTR)sPath );
									}
								}
								else
								{
									sResult = ErrorMessage( hres );
									TRACE( _T("GetPath() error %s : %s\n"), (LPCTSTR)sResult, (LPCTSTR)sPath );
								}
							}
							else
							{
								sResult = ErrorMessage( hres );
								TRACE( _T("Shell link error %s : %s\n"), (LPCTSTR)sResult, (LPCTSTR)sPath );
							}
						}
						else
						{
							sResult = ErrorMessage( hres );
							TRACE( _T("Failed CoCreateInstance() %s : %s\n"), (LPCTSTR)sResult, (LPCTSTR)sPath );
						}

						if ( nType != LinkType::Unknown )
						{
							SHFILEINFO sfi = {};
							SHGetFileInfo( sPath, 0, &sfi, sizeof( sfi ), SHGFI_ICON | SHGFI_SMALLICON );
							OnNewItem( new CLink( nType, sfi.hIcon, sPath, sTarget, sResult, bResult ) );
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

	PostMessage( WM_TIMER, ID_DONE );
}
