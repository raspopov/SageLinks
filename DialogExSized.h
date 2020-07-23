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

// CDialogExSized dialog

class CDialogExSized : public CDialogEx
{
	DECLARE_DYNAMIC(CDialogExSized)

public:
	CDialogExSized() = default;
	CDialogExSized(UINT nIDTemplate, CWnd *pParent = NULL) : CDialogEx( nIDTemplate, pParent ) {}
	CDialogExSized(LPCTSTR lpszTemplateName, CWnd *pParentWnd = NULL) : CDialogEx( lpszTemplateName, pParentWnd ) {}
	virtual ~CDialogExSized() override = default;

	void ReloadLayout();

	void SaveWindowPlacement();
	void RestoreWindowPlacement();

private:
	CRect m_rcInitial,  m_rcInitialClient;		// Начальный (минимальный) размер окна

protected:
	virtual BOOL OnInitDialog() override;

	afx_msg void OnGetMinMaxInfo( MINMAXINFO* lpMMI );
	afx_msg void OnDestroy();

	DECLARE_MESSAGE_MAP()
};
