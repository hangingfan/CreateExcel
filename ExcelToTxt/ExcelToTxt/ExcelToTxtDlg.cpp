// ExcelToTxtDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "ExcelToTxt.h"
#include "ExcelToTxtDlg.h"

#include <iostream>
#include <windows.h>
#include "include_cpp\libxl.h"
#include <map>
#include <cstdlib>
#include <sstream>
#include <algorithm>
#include <vector>
#include <fstream>
#include <stdlib.h> 
#include "direct.h" 
#include "shlobj.h"  

#pragma comment(lib, "libxl.lib")
bool m_bForSubGalaxy = false;


using namespace libxl;
using namespace std;

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


const wstring m_TagStr = L"&&";
Book* book;
Sheet* sheet;
int m_iMaxRow = 0;               //�����������
wstring strContent = L"";
string strSavePath = "";        //�ļ������·��
bool  m_b10000ToFloat = false;
vector<int> m_10000ToFloatVec;     //����Ҫ�������ת������

// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CExcelToTxtDlg �Ի���

std::string ws2s(const std::wstring& ws)
{
	std::string curLocale = setlocale(LC_ALL, NULL);        // curLocale = "C";
	setlocale(LC_ALL, "chs");
	const wchar_t* _Source = ws.c_str();
	size_t _Dsize = 2 * ws.size() + 1;
	char *_Dest = new char[_Dsize];
	memset(_Dest, 0, _Dsize);
	wcstombs(_Dest, _Source, _Dsize);
	std::string result = _Dest;
	delete[]_Dest;
	setlocale(LC_ALL, curLocale.c_str());
	return result;
}


//��ȡ�����ļ��еı���·��
string ReadSavePath()
{
		TCHAR currentDir[MAX_PATH];
		GetCurrentDirectory( MAX_PATH, currentDir );
		wstring test(&currentDir[0]); //convert to wstring
		string strFolder(test.begin(), test.end());

		FILE *file;
		strFolder = strFolder + "/Settings.txt";
		fopen_s(&file, strFolder.c_str(), "r, ccs=UTF-8");
		if(file == NULL)
		{
			return "";
		}

		//������ǿ��ַ�
		wchar_t buffer[MAX_PATH] =L"";
		fread( buffer,  sizeof( wchar_t ),MAX_PATH, file );
		fclose(file); 

		wstring wsContent(buffer);
		string strTest(wsContent.begin(),wsContent.end());

		return strTest;
}


CExcelToTxtDlg::CExcelToTxtDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CExcelToTxtDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CExcelToTxtDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CExcelToTxtDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_DROPFILES() 

	//}}AFX_MSG_MAP
	//ON_MESSAGE(WM_DROPFILES,OnDropFiles)
	ON_BN_CLICKED(IDC_BUTTON_SETTINGS, &CExcelToTxtDlg::OnBnClickedButtonSettings)
END_MESSAGE_MAP()


// CExcelToTxtDlg ��Ϣ�������

BOOL CExcelToTxtDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// To Accept Dropped file Set this TRUE
	DragAcceptFiles(TRUE);

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

wstring IntToWString(int  iNum)
{
	std::wostringstream ws;
	ws << iNum;
	wstring strInt(ws.str());
	return strInt;
}


//ר��Ϊ�����ǽ������������ɵ�ͼƬID,��0-20������ɡ��ؼ�����SubName5
wstring AnalysisForSubGalaxy()
{
	int iFirst = 0;
	int iLast = 20;

	int iPicSerial0 = rand() % iLast + iFirst;
	int iPicSerial1 = rand() % iLast + iFirst;
	while (iPicSerial1 == iPicSerial0)
	{
		iPicSerial1 = rand() % iLast + iFirst;
	}

	int iPicSerial2 = rand() % iLast + iFirst;
	while (iPicSerial2 == iPicSerial0 || iPicSerial2 == iPicSerial1)
	{
		iPicSerial2 = rand() % iLast + iFirst;
	}

	int iPicSerial3 = rand() % iLast + iFirst;
	while (iPicSerial3 == iPicSerial0 || iPicSerial3 == iPicSerial1 || iPicSerial3 == iPicSerial2)
	{
		iPicSerial3 = rand() % iLast + iFirst;
	}

	int iPicSerial4 = rand() % iLast + iFirst;
	while (iPicSerial4 == iPicSerial0 || iPicSerial4 == iPicSerial1 || iPicSerial4 == iPicSerial2 || iPicSerial4 == iPicSerial3)
	{
		iPicSerial4 = rand() % iLast + iFirst;
	}

	wstring strTemp = m_TagStr + IntToWString(iPicSerial0) + m_TagStr + IntToWString(iPicSerial1) + m_TagStr + IntToWString(iPicSerial2) + m_TagStr + IntToWString(iPicSerial3) + m_TagStr + IntToWString(iPicSerial4);

	return strTemp;
}



void Analysis()
{
	//�ȱ�����һ�У��ҵ����в���Ҫ�������ת������
	int i = 0;
	while (true)
	{
		int iError = sheet->cellType(0, i);
		if (iError == CELLTYPE_EMPTY)
		{
			m_iMaxRow = i;
			break;
		}

		if (iError == CELLTYPE_STRING)
		{
			const wchar_t* Content = sheet->readStr(0, i);
			wstring str(Content);

			if(str.find(L"����Ʊ��") != wstring::npos)
			{
				m_b10000ToFloat = true;
				m_10000ToFloatVec.push_back(i);
			}
		}
		++i;
	}

	i = 0;
	//�ҵ��������Ƶ���һ�У����û���ҵ���˵����û���������ƣ��Ͳ���ɾ�����һ��
	while (true)
	{
		int iError = sheet->cellType(0, i);
		if (iError == CELLTYPE_EMPTY)
		{
			m_iMaxRow = i;
			break;
		}

		if (iError == CELLTYPE_STRING)
		{
			const wchar_t* Content = sheet->readStr(0, i);
			wstring str(Content);

			//ר��Ϊ�����ǽ������������ɵ�ͼƬID,��0-20������ɡ��ؼ�����SubName5
			if (str == L"SubName5ForSubGalaxy")
			{
				m_bForSubGalaxy = true;
			}

			if (str == L"NumberLimit")
			{
				m_iMaxRow = i;
				break;
			}
		}

		++i;
	}

	//�ӵڶ��п�ʼ��ȡ���ݣ���һ����˵��������Ҫ
	for (int row = 1; sheet->cellType(row, 0) != CELLTYPE_EMPTY; ++row)
	{
		for (int line = 0; line < m_iMaxRow; ++line)
		{
			if (sheet->cellType(row, line) == CELLTYPE_STRING)
			{
				wstring str = sheet->readStr(row, line);
				if(str == L"isempty")   //������ַ����� ��û������
				{
					strContent = strContent + m_TagStr;
				}
				else
				{
					strContent = strContent + str + m_TagStr;
				}
			}
			else if (sheet->cellType(row, line) == CELLTYPE_NUMBER)
			{
				long iNum = sheet->readNum(row, line);

				double dNum = sheet->readNum(row, line);   //����excel��4λС����
				char test[100];
			    sprintf(test, "%.4f", dNum);
				string strRet(test);
				float fNum = ::atof(strRet.c_str()); 
				float fTemp = (float)iNum;
				if (fNum == fTemp)
				{
					bool bRet = false;
					if(m_b10000ToFloat && iNum != -1)  //��Ҫ������ֱ仯����
					{
						if(find(m_10000ToFloatVec.begin(), m_10000ToFloatVec.end(), line) == m_10000ToFloatVec.end())
						{
							bRet = true;
						}
					}

					if(bRet)
					{
						char IntToFloat[100];    //������������У�ת��ΪС��λ
						double NewNum = dNum/10000;
						sprintf(IntToFloat, "%.4f", NewNum);
						string NewStrRet(IntToFloat);

						wstring strFloat(NewStrRet.begin(), NewStrRet.end());
						strContent = strContent + strFloat + m_TagStr;
					}
					else
					{
						wstring strInt = IntToWString(iNum);
						strContent = strContent + strInt + m_TagStr;
					}
				}
				else
				{
					if(dNum > 1000000 && iNum < 0)   //������Long�����ֵ
					{
						char Limittest[100];
						sprintf(Limittest, "%.0f", dNum);
						string NewStrRet(Limittest);

						wstring strFloat(NewStrRet.begin(), NewStrRet.end());
						strContent = strContent + strFloat + m_TagStr;
					}
					else
					{
						wstring strFloat(strRet.begin(), strRet.end());
						strContent = strContent + strFloat + m_TagStr;
					}
					
				}
			}
			else if (sheet->cellType(row, line) == CELLTYPE_EMPTY)
			{
				strContent += m_TagStr;
			}
		}

		strContent = strContent.substr(0, strContent.length() - 2);

		if (m_bForSubGalaxy)
		{
			strContent += AnalysisForSubGalaxy();
		}

		strContent += L"\n";
	}
}

//��ÿһ���ļ����ж�ȡ
void ProcessPerFile(CString fileName)
{
	int length = fileName.GetLength();
	CString CSTemp = fileName.Mid(length - 4, 4);
	if(CSTemp != ".xls")
	{
		MessageBoxW(NULL, _T("��ѡ��xls�ļ��� ���ð�!"), _T("Error"), MB_ICONERROR | MB_OK);
		return;
	}

	m_b10000ToFloat = false;
	strContent = L"";
	wstring strPath(fileName);

	//������������
	book = xlCreateBook();
	if (book)
	{
		///****************   ��ȡ����   ****************/
		if (book->load(strPath.c_str()))
		{
			int iSerial = 0;
			do 
			{
				sheet = book->getSheet(iSerial);
				if (sheet == NULL)
				{
					break;
				}

				Analysis();

				m_iMaxRow = 0;
				++iSerial;
			} while (true);


			
			if (iSerial != 0)
			{
				strContent = strContent.substr(0, strContent.length() - 1);

				ofstream myfile;

				strPath.replace(strPath.length() - 3, strPath.length(), L"txt");
				string strTest = ws2s(strPath);

				int iResultSerial = strTest.rfind("\\");
				//string strResultDir = strTest.substr(0, iResultSerial) + "\\Result";  //�洢Result�����ļ���
				//_mkdir(strResultDir.c_str());
			 //   

				strTest = strSavePath + strTest.substr(iResultSerial, strTest.length() - iResultSerial);
				//const WCHAR * wpszProcToSearch = strPath.c_str();

				FILE *file;
				fopen_s(&file, strTest.c_str(), "w, ccs=UTF-8");
				fwrite(strContent.c_str(), sizeof(wchar_t), strContent.length(), file);
				fclose(file);
				
			}
		}

		
	}
	book->release();
}

void CExcelToTxtDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CExcelToTxtDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CExcelToTxtDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CExcelToTxtDlg::OnDropFiles (HDROP dropInfo) 
{ 
    CString sFile; 
    DWORD   nBuffer = 0; 

	strSavePath = ReadSavePath();
	if(strSavePath == "")
	{
		MessageBoxW( _T("���������ļ�����·��"), _T("Error"), MB_ICONERROR | MB_OK);
		return;
	}
 
    // Get the number of files dropped 
    int nFilesDropped = DragQueryFile (dropInfo, 0xFFFFFFFF, NULL, 0); 
 
    for(int i=0; i<nFilesDropped; i++) 
    { 
        // Get the buffer size of the file. 
        nBuffer = DragQueryFile (dropInfo, i, NULL, 0); 
 
        // Get path and name of the file 
        DragQueryFile (dropInfo, i, sFile.GetBuffer (nBuffer + 1), nBuffer + 1); 
     
        //�������еĲ���
		m_bForSubGalaxy = false;
		m_10000ToFloatVec.clear();

        m_iMaxRow = 0;           //�����������
		wstring strContent = L"";

		ProcessPerFile(sFile.GetBuffer());
		sFile.ReleaseBuffer (); 
    } 

	if(nFilesDropped != 0)
	{
		MessageBoxW( _T("�ļ�ת�����"), _T("Success"), MB_OK);
		
	}

	//for(int i = 0; i < 
 
    // Free the memory block containing the dropped-file information 
    DragFinish(dropInfo); 
} 

bool GetFolder(std::string& folderpath, const char* szCaption = NULL, HWND hOwner = NULL)    
{    
    bool retVal = false;    
    
    // The BROWSEINFO struct tells the shell    
    // how it should display the dialog.    
    BROWSEINFO bi;    
    memset(&bi, 0, sizeof(bi));    
    bi.ulFlags   = BIF_USENEWUI;    
    bi.hwndOwner = hOwner;    
    bi.lpszTitle = (LPCWSTR)szCaption;    
    
    // must call this if using BIF_USENEWUI    
    ::OleInitialize(NULL);    
    
    // Show the dialog and get the itemIDList for the selected folder.    
    LPITEMIDLIST pIDL = ::SHBrowseForFolder(&bi);    
    
    if(pIDL != NULL)    
    {    
        // Create a buffer to store the path, then get the path.    
        char buffer[_MAX_PATH] = {'\0'};    
		LPWSTR wideStr = new TCHAR[_MAX_PATH];
        if(::SHGetPathFromIDList(pIDL, wideStr) != 0)    
        {    
			wcstombs(buffer, wideStr, _MAX_PATH);
            // Set the string value.    
            folderpath = buffer;    
            retVal = true;    
        }           
    
        // free the item id list    
        CoTaskMemFree(pIDL);    
    }    
    
    ::OleUninitialize();    
    
    return retVal;    
}   

//�������ɵ��ļ�·��
void CExcelToTxtDlg::OnBnClickedButtonSettings()
{
	std::string szPathContent("");    
    
	if (GetFolder(szPathContent, "Select a folder.") == true)    
	{    
		TCHAR currentDir[MAX_PATH];
		GetCurrentDirectory( MAX_PATH, currentDir );
		wstring test(&currentDir[0]); //convert to wstring
		string strFolder(test.begin(), test.end());

		wstring wsContent(szPathContent.begin(), szPathContent.end());
		FILE *file;
		strFolder = strFolder + "/Settings.txt";
		fopen_s(&file, strFolder.c_str(), "w, ccs=UTF-8");
		fwrite(wsContent.c_str(), sizeof(wchar_t), wsContent.length(), file);
		fclose(file);  
	}    
}
