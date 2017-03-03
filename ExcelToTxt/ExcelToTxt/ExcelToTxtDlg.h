// ExcelToTxtDlg.h : 头文件
//

#pragma once


// CExcelToTxtDlg 对话框
class CExcelToTxtDlg : public CDialog
{
// 构造
public:
	CExcelToTxtDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_EXCELTOTXT_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnDropFiles (HDROP dropInfo); 
	DECLARE_MESSAGE_MAP()
public:
	CListBox list; 
	afx_msg void OnBnClickedButtonSettings();
	
};
