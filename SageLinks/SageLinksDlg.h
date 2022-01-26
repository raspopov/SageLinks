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

#include "DialogExSized.h"

// CSageLinksDlg dialog

class CSageLinksDlg : public CDialogExSized
{
	DECLARE_DYNAMIC(CSageLinksDlg)

public:
	CSageLinksDlg(CWnd* pParent = nullptr);

// Dialog Data
	enum { IDD = IDD_SAGELINKS_DIALOG };

	afx_msg BOOL OnKeyDown( UINT nChar, UINT nRepCnt, UINT nFlags );

	CString				m_sPath;			// Initial path

protected:
	enum LinkType { Unknown = 0, Symbolic, Junction, Shortcut, AppX };

	class CLink
	{
	public:
		inline CLink(LinkType nType, HICON hIcon, const CString& sSource, const CString& sTarget, const CString& sResult, BOOL bResult) noexcept
			: m_nType	( nType )
			, m_hIcon	( hIcon )
			, m_sSource	( sSource )
			, m_sTarget	( sTarget )
			, m_sResult	( sResult )
			, m_bResult	( bResult )
		{
		}

		CLink(CLink&) = delete;
		CLink& operator=(CLink&) = delete;

		inline ~CLink() noexcept
		{
			if ( m_hIcon )
			{
				DestroyIcon( m_hIcon );
			}
		}

		inline bool IsExist() const noexcept
		{
			return ::IsExist( m_sSource );
		}

		inline bool DeleteFile() const noexcept
		{
			return ::DeleteFile( IsUNC( m_sSource ) ? m_sSource : ( LONG_PREFIX + m_sSource ) );
		}

		LinkType	m_nType;
		HICON		m_hIcon;
		CString		m_sSource;
		CString		m_sTarget;
		CString		m_sResult;
		BOOL		m_bResult;
	};

	using CLinkList = std::vector< CLink* >;

	HICON				m_hIcon;
	CImageList			m_oImages;
	int					m_nImageSuccess;
	int					m_nImageError;
	int					m_nImageType[ 5 ];

	int					m_nBad;				// Bad links
	int					m_nSort;			// Sort column
	CListCtrl			m_wndList;
	CStatic				m_wndStatus;
	CStringList			m_oDirs;
	CString				m_sStatus;
	CString				m_sOldStatus;
	HANDLE				m_hThread;
	CCriticalSection	m_pSection;
	CLinkList			m_pList;			// Incoming items
	CEvent				m_pFlag;
	CMFCEditBrowseCtrl	m_wndBrowse;

	static unsigned __stdcall ThreadStub(void* param);
	void Thread();
	void Stop();
	void Start();

	void DeSelect();
	void SortList();
	void ClearList();
	void OnNewItem(CLink* link);
	void Resize();

	virtual BOOL OnInitDialog() override;
	virtual void DoDataExchange( CDataExchange* pDX ) override;
	virtual void OnOK() override;

	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	afx_msg void OnDestroy();
	afx_msg void OnSize( UINT nType, int cx, int cy );
	afx_msg void OnNMDblclkList( NMHDR *pNMHDR, LRESULT *pResult );
	afx_msg void OnNMRClickList( NMHDR *pNMHDR, LRESULT *pResult );
	afx_msg void OnLvnGetdispinfoList( NMHDR *pNMHDR, LRESULT *pResult );
	afx_msg void OnLvnOdcachehintList( NMHDR *pNMHDR, LRESULT *pResult );
	afx_msg void OnLvnOdfinditemList( NMHDR *pNMHDR, LRESULT *pResult );
	afx_msg void OnHdnItemclickList( NMHDR *pNMHDR, LRESULT *pResult );
	afx_msg void OnBnClickedDelete();

	DECLARE_MESSAGE_MAP()
};

#define ID_TIMER	100		// Timer notification message
#define ID_DONE		101		// End of work notification message
