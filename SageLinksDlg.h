/*
This file is part of SageLinks

Copyright (C) 2015-2017 Nikolay Raspopov <raspopov@cherubicsoft.com>

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


// CSageLinksDlg dialog

class CSageLinksDlg : public CDialog
{
public:
	CSageLinksDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_SAGELINKS_DIALOG };
#endif

	afx_msg BOOL OnKeyDown( UINT nChar, UINT nRepCnt, UINT nFlags );

protected:
	enum LinkType { Unknown, Symbolic, Junction, Shortcut };

	class ATL_NO_VTABLE CLink
	{
	public:
		inline CLink(LinkType nType = LinkType::Unknown, HICON hIcon = NULL, const CString& sSource = CString(), const CString& sTarget = CString(), BOOL bResult = FALSE)
			: m_nType	( nType )
			, m_hIcon	( hIcon )
			, m_sSource	( sSource )
			, m_sTarget	( sTarget )
			, m_bResult	( bResult )
		{
		}
		
		inline ~CLink()
		{
			if ( m_hIcon ) DestroyIcon( m_hIcon );
		}

		LinkType	m_nType;
		HICON		m_hIcon;
		CString		m_sSource;
		CString		m_sTarget;
		BOOL		m_bResult;
	};
	typedef CList< CLink* > CLinkList;

	HICON				m_hIcon;
	CImageList			m_oImages;
	int					m_nImageSuccess;
	int					m_nImageError;
	int					m_nImageUnknown;
	int					m_nImageSymbolic;
	int					m_nImageJunction;
	int					m_nImageShortcut;

	int					m_nTotal;			// Total links
	int					m_nBad;				// Bad links
	
	BOOL				m_bSort;			// Sort needed
	CListCtrl			m_wndList;
	CStatic				m_wndStatus;
	CStringList			m_oDirs;
	CString				m_sStatus;
	CString				m_sOldStatus;
	HANDLE				m_hThread;
	CCriticalSection	m_pSection;
	CLinkList			m_pIncoming;		// Incoming items
	CEvent				m_pFlag;
	CMFCEditBrowseCtrl	m_wndBrowse;
	CRect				m_rcInitial;

	static unsigned __stdcall ThreadStub(void* param);
	void Thread();
	void Stop();
	void Start();

	void SortList();
	static int CALLBACK SortFunc( LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort );
	void ClearList();
	void OnNewItem(CLink* pLink);

	virtual BOOL OnInitDialog();
	virtual void DoDataExchange( CDataExchange* pDX );	// DDX/DDV support
	virtual void OnOK();

	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnTimer(UINT_PTR nIDEvent);
	afx_msg void OnDestroy();
	afx_msg void OnSize( UINT nType, int cx, int cy );
	afx_msg void OnGetMinMaxInfo( MINMAXINFO* lpMMI );
	afx_msg void OnNMDblclkList( NMHDR *pNMHDR, LRESULT *pResult );
	afx_msg void OnNMRClickList( NMHDR *pNMHDR, LRESULT *pResult );
	afx_msg void OnHdnItemclickList( NMHDR *pNMHDR, LRESULT *pResult );

	DECLARE_MESSAGE_MAP()
};
