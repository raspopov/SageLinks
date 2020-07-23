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

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols


// CSageLinksApp

class CSageLinksApp : public CWinApp
{
public:
	CSageLinksApp();

protected:
	virtual BOOL InitInstance();
	virtual BOOL ProcessMessageFilter(int code, LPMSG lpMsg);

	DECLARE_MESSAGE_MAP()
};

extern CSageLinksApp theApp;

#define OPT_SECTION		_T("Settings")
#define OPT_PATH		_T("Path")

#define TAG_PREFIX		_T("\\??\\")
#define UNC_PREFIX		_T("\\??\\UNC\\")
#define LONG_PREFIX		_T("\\\\?\\")

#define LONG_PATH		512	// symbols

#define _P(x)			(x),(_countof(x)-1)

CString ErrorMessage(HRESULT hr);

inline BOOL IsExist(LPCTSTR szPath) noexcept
{
	return ( ::GetFileAttributes( szPath ) != INVALID_FILE_ATTRIBUTES );
}

inline BOOL IsDots(LPCTSTR szFileName) noexcept
{
	return szFileName[ 0 ] == _T('.') && ( szFileName[ 1 ] == 0 || ( szFileName[ 1 ] == _T('.') && szFileName[ 2 ] == 0 ) );
}

inline BOOL IsLocal(LPCTSTR szFileName) noexcept
{
	return _istalpha( szFileName[ 0 ] ) && szFileName[ 1 ] == _T( ':' ) && ( szFileName[ 2 ] == 0 || szFileName[ 2 ] == _T('\\') );
}

inline BOOL IsUNC(LPCTSTR szFileName) noexcept
{
	return szFileName[ 0 ] == _T('\\') && szFileName[ 1 ] == _T('\\') && _istalnum( szFileName[ 2 ] );
}

inline CString GetRemotedPath(const CString& sBasePath, const CString& sLocalPath)
{
	// Source path is UNC
	const int n = sBasePath.Mid( 2 ).Find( _T( '\\' ) ) + 2;
	if ( n > 2 && IsLocal( sLocalPath ) )
	{
		// Target path is local and source path is rooted UNC
		return sBasePath.Left( n + 1 ) + sLocalPath.GetAt( 0 ) + _T( '$' ) + sLocalPath.Mid( 2 );
	}
	return CString();
}
