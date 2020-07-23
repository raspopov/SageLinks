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

#include "Shell.h"

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
				TPM_LEFTBUTTON | TPM_RIGHTBUTTON, point.x, point.y, 0, hwnd, nullptr );
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
