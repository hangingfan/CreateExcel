// ExcelToTxtDlg.h : ͷ�ļ�
//

#pragma once


// CExcelToTxtDlg �Ի���
class CExcelToTxtDlg : public CDialog
{
// ����
public:
	CExcelToTxtDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_EXCELTOTXT_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
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
