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

class CShellItem
{
public:
	CShellItem(LPCTSTR szFullPath, void** ppFolder = nullptr, HWND hWnd = nullptr)
		: m_pidl	( nullptr )
		, m_pLastId	( nullptr )
	{
		CComPtr< IShellFolder > pDesktop;
		HRESULT hr = SHGetDesktopFolder( &pDesktop );
		if ( FAILED( hr ) )
			return;

		hr = pDesktop->ParseDisplayName( hWnd, nullptr, CT2OLE( szFullPath ), nullptr, &m_pidl, nullptr );
		if ( FAILED( hr ) )
			return;

		m_pLastId = ILFindLastID( m_pidl );

		if ( ppFolder )
		{
			USHORT temp = m_pLastId->mkid.cb;
			m_pLastId->mkid.cb = 0;
			hr = pDesktop->BindToObject( m_pidl, nullptr, IID_IShellFolder, ppFolder );
			m_pLastId->mkid.cb = temp;
		}
	}

	~CShellItem() noexcept
	{
		if ( m_pidl )
		{
			CoTaskMemFree( m_pidl );
		}
	}

	inline operator LPITEMIDLIST() const noexcept
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
		: m_pID	( nullptr )
	{
		for ( POSITION pos = oFiles.GetHeadPosition(); pos; )
		{
			const CString strPath = oFiles.GetNext( pos );

			CShellItem* pItemIDList = new (std::nothrow) CShellItem( strPath, ( m_pFolder ? nullptr : (void**)&m_pFolder ) );	// Get only one
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

		m_pID.Attach( new (std::nothrow) PCUITEMID_CHILD[ GetCount() ] );
		if ( ! m_pID )
			// Out of memory
			return;

		size_t i = 0;
		for ( POSITION pos = GetHeadPosition(); pos; ++ i )
			m_pID[ i ] = GetNext( pos )->m_pLastId;
	}

	virtual ~CShellList() noexcept
	{
		for ( POSITION pos = GetHeadPosition(); pos; )
		{
			delete GetNext( pos );
		}
		RemoveAll();
	}

	// Creates menu from file paths list
	inline bool GetMenu(HWND hWnd, void** ppContextMenu) noexcept
	{
		return m_pFolder && SUCCEEDED( m_pFolder->GetUIObjectOf( hWnd, (UINT)GetCount(), m_pID, IID_IContextMenu, nullptr, ppContextMenu ) );
	}

protected:
	CComPtr< IShellFolder >				m_pFolder;	// First file folder
	CAutoVectorPtr< PCUITEMID_CHILD >	m_pID;		// File ItemID array
};
