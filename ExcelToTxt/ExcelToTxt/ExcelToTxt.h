// ExcelToTxt.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CExcelToTxtApp:
// �йش����ʵ�֣������ ExcelToTxt.cpp
//

class CExcelToTxtApp : public CWinApp
{
public:
	CExcelToTxtApp();

// ��д
	public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CExcelToTxtApp theApp;