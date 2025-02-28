/*
 * MFCExcel.h
 * Author: hearstzhang (hearstzhang@tencent.com)
 * A safely MFCExcelFile which added many useful functions
 * and mechanisms to make it much faster compared to the gyssoft's
 * original version of MFCExcelFile.
 * It will finish read and write for millions of cells in seconds, if you use it well.
 * Preload excel sheet is recommended, unless memory problem.
 * Your permission for copy and redistribute is granted as long as you keep
 * these copyright announcements.
 * Requires Microsoft Excel or Kingsoft WPS installed. Compatible with
 * MBCS and Unicode character set.
 * Please let me know if you have some question about this.
 */

/*
 * v1.0.0 initial version, compatible to both MBCS and UNICODE
 * v1.0.1 add two batch set cell string function
 * v1.0.2 fixed bugs for pretential crash at preload mode
 * v1.0.3 fixed bugs for set cell string function
 * v1.0.4 fixed bugs for pretential crash at init and release excel application
 * v1.0.5 add two batch set cell string function, which could clear rest rows and cols
 * v1.0.6 add exception handlers for multiple function which could make it stable
 * v1.0.7 fixed bugs for batch set cell string function.
 * v1.0.8 increase performance of batch set cell string function.
 * v1.0.9 fixed bugs for pretential crash.
 * v1.0.10 add delete row and col function
 * v1.0.11 delete row and col function does not require preload
 * v1.0.12 batch set cell and delete rest row and col function does not require preload
 * v1.0.13 fixed bugs for delete row and col function
 * v1.0.14 add get cell name function
 * v1.0.15 increase performance of batch set cell string
 * v1.0.16 batch set cell string does not do nothing if vector is not big enough.fill rest cells as blank instead.
 * v1.0.17 get cell name function does not requries preload.
 * v1.0.18 increase performance for reading for preload mode
 * v1.0.19 increase performance for set cell value one by one.
 * v1.0.20 loadsheet function could automatically create a blank sheet if target sheet does not exist.
 * v1.0.21 openexcelfile function could automatically create a excel file in memory for operation if target excel file does not exist.But this file will save to disk only if you save it.
 * v1.0.22 fixed a bug which delete ranged row may not take effect in some circumstances.
 * v1.0.23 add save function, could save the excel to disk and reload it immediately to refresh status.
 * v1.0.24 move delete xls method enum to source file to avoid global variable define.
 * v1.0.25 add singleton adminstration class llusionExcelSingletonAdmin for MFCExcelFile.But compatible with old codes.
 * v1.0.26 make modification to MFCExcelFile class constructor and destructor to compatible with llusionExcelSingletonAdmin class.
 * v1.0.27 fix a bug which may cause argument error if preload an empty sheet.
 * v1.0.28 PreloadSheet is a friend function now, which means, no allowed to use outside the class.
 * v1.0.29 Solve pretential memory leak problem.
 * v1.0.30 Preload mode does not require save function to refresh memory data.This option will be done by any read functions automatically.
 * v1.0.31 Update copyright information.
 * v1.0.32 Support absolate and not absolate path.
 * v1.0.33 Add demo excel operation function. See MFCExcel.cpp for details.
 */

// -----------------------MFCExcel.h------------------------
#pragma once

#include "CApplication.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"
#include "CRange.h"
#include <vector>

// ǰ������
class MFCExcelFile;

// ���ڹ���MFCExcel�ĵ�����
// ������ʹ���������������MFCExcelFile�ࡣ������Ҫע�⣬����౾���ǵ���ģʽ���࣬
// �����ü���Ϊ0��ʱ���Զ�������MFCExcelFile����
// ��֤��Ĺ������˳���ʱ�򣬿�����ȷ�ͷŵ�Excel App
class MFCExcelSingletonAdmin
{
protected:
	// MFCExcelFile��������
	static MFCExcelFile * m_pInstance;
	// ���ü�����
	static unsigned int m_nReference;

public:
	// �������������
	// ��Ϊpublic����Ϊ����౾���ǵ���ģʽ����
	MFCExcelSingletonAdmin();
	~MFCExcelSingletonAdmin();
	// ���Ŀ������û�а�װExcel��WPS�򷵻�nullptr
	// ���򷵻�һ��������MFCExcelFile����
	MFCExcelFile * GetInstance();
};

// ����ʹ�����ĵ���ģʽ������MFCExcelSingletonAdmin����
// ��Ϊһ������ֻ�ܵ���һ��excel����
// Ҳ����˵����������ȡ��һ��excel�ļ������Ҫ�رյ�ǰ�ļ���������ReleaseExcel��
// index��1�����Ǵ�0��ʼ������col = 1, row = 1��Ӧ��Ԫ��A1��sheetindex = 1��Ӧ��һ��sheet�ȡ�
class MFCExcelFile
{
	// ����ģʽ��������Ԫ
	friend class MFCExcelSingletonAdmin;
protected:
	// Ԥ���غ���
	// ��Ԥ����ģʽ�£��κζ�Excel��д�붼������һ�ζ�ȡ���ݵĲ�����ʱ��
	// �Զ�������ˢ���ڴ渱����
	void PreLoadSheet();
	// �Զ������·��ת���ɾ���·��
	// �Զ���\ת����//�Ժ�MFC���ݡ�
	// ���������Ǿ���·�����򲻻��path���κθı䣬����ת��\
	// ����������pathת���ɾ���·�����Ե�ǰ��������·��Ϊ��ʼ
	void GetRelativePathIfNot(CString &path);

public:
	// ���캯������������
	MFCExcelFile();
	virtual ~MFCExcelFile();

protected:
	// �򿪵�EXCEL�ļ�����
	CString open_excel_file_;
	// EXCEL BOOK���ϣ�������ļ�ʱ��
	CWorkbooks excel_books_;
	// ��ǰʹ�õ�BOOK����ǰ������ļ�
	CWorkbook excel_work_book_;
	// EXCEL��sheets����
	CWorksheets excel_sheets_;
	// ��ǰʹ��sheet
	CWorksheet excel_work_sheet_;
	// ��ǰ�Ĳ�������
	CRange excel_current_range_;
	// �Ƿ��Ѿ�Ԥ������ĳ��sheet������
	BOOL already_preload_;
	std::vector<int> a;
	// Ԥ����ģʽ��ʹ��OLE����洢Excel����
	COleSafeArray ole_safe_array_;
	// ά��Excel�����Ƿ��Ѿ�����ʼ�������ͷš�
	static bool isInited;
	static bool isReleased;
	// EXCEL�Ľ���ʵ��
	static CApplication excel_application_;
	// Ԥ����ģʽ�£���������Ƿ���Ҫˢ��Ԥ�������ݣ�Ϊtrue����Ҫˢ��
	bool isUpdate;

public:
	// ����Excel�򿪵�ǰ�ĵ�
	void ShowInExcel(BOOL bShow = TRUE);
	// ���һ��CELL�Ƿ����ַ���
	BOOL IsCellString(long iRow, long iColumn);
	// ���һ��CELL�Ƿ�����ֵ
	BOOL IsCellInt(long iRow, long iColumn);
	// �õ�һ��CELL��String�������Ԥ���ص�����£����Է���Խ�����ݣ���᷵�ؿհ�ֵ��
	CString GetCellString(long iRow, long iColumn);
	// �õ������������Ԥ���ص�����£����Է���Խ�����ݣ���᷵��0
	int GetCellInt(long iRow, long iColumn);
	// �õ�double�����ݣ������Ԥ���ص�����£����Է���Խ�����ݣ���᷵��0.0
	double GetCellDouble(long iRow, long iColumn);
	// ȡ���е�����
	long GetRowCount();
	// ȡ���е�����
	long GetColumnCount();
	// ����Sheet�Թ�ʹ�ã����û�У����Զ�����һ���������¼���Sheet��Ÿ���
	// ��ǰ���е�Sheet����������Զ�����󴴽�Sheetֱ�������������index��
	// ע��index��1��ʼ���������0��ֱ�ӷ���FALSE.
	// PreloadΪTRUE��Ԥ����ģʽ�����������ȡ���ݣ�������ʹ�ø�ģʽ������Sheet�ǳ���ʹ��Ԥ���س������ļ������Դ
	BOOL LoadSheet(long table_index, BOOL pre_load = TRUE);
	// ͨ������ʹ��ĳ��sheet��
	// ���û���������������Sheet����������ṩ�������Զ�����һ����������档
	// Preload��Ԥ����ģʽ�����������ȡ���ݣ�������ʹ�ø�ģʽ������Sheet�ǳ���ʹ��Ԥ���س������ļ������Դ��
	BOOL LoadSheet(const TCHAR* sheet, BOOL pre_load = TRUE);
	// ͨ�����ȡ��ĳ��Sheet������
	CString GetSheetName(long table_index);
	// �õ�Sheet������
	long GetSheetCount();
	// ���ļ��������Ӧλ��û�����excel�ļ������Զ����ڴ洴��һ���հ׵�excel�ļ���������Ч������FALSE���ʧ��
	// ֧�־��Ժ����·������������·�����������������ڵĵ�ǰ·��Ϊ��ʼ��
	// ֧����б�ܺͷ�б�ܵ�·��������"..\\1.xlsx"��"../1.xlsx"��
	BOOL OpenExcelFile(const TCHAR * file_name);
	// �رմ򿪵�Excel �ļ������if_saveΪTRUE�򱣴��ļ�
	void CloseExcelFile(BOOL if_save = FALSE);
	// ���Ϊһ��EXCEL�ļ�
	// ֧�־��Ժ����·������������·�����������������ڵĵ�ǰ·��Ϊ��ʼ��
	// ֧����б�ܺͷ�б�ܵ�·��������"..\\1.xlsx"��"../1.xlsx"��
	void SaveasXLSFile(const CString &xls_file);
	// ���̱��浽���̣�ͬʱ��ر��ٴ򿪵�ǰsheet��ˢ��״̬��
	bool GetRangeCellString(DWORD iRowStart, DWORD iRowEnd, DWORD iColStart, DWORD iColEnd,
		std::vector<CString>& new_string);
	void Save();
	// ȡ�ô��ļ�������
	CString GetOpenedFileName();
	// ȡ�ô�sheet������
	CString GetLoadSheetName();
	// д��һ��CELLһ��int�������ܲ�Ҫ��������д����Ϊ�ú������۽ϸ�
	void SetCellInt(long irow, long icolumn, int new_int);
	// д��һ��CELLһ��string�������ܲ�Ҫ��������д����Ϊ�ú������۽ϸ�
	void SetCellString(long irow, long icolumn, CString new_string);
	// д��һ����Χ��string�������A2д��D10������2, 10, 1, 4�����д����vector����洢���ݣ����ұ�����ʽ��new_string[irow * colsize + icol]
	// ���iRowStart <= 0��iColStart <= 0�����������ʲô������ֱ�ӷ���false��
	// vector������û��ϵ���Ὣ������Χ������д�ɿհף�����ҪС������㷶Χ���õĲ���ȷ����ÿհ״���ĸ��ǵ���Ӧ���ݡ�
	// ������൱��Ч�ĺ���������˵��д��ǧ�����������ݣ���i7-4790��16GDDR3 32λExcel 2016��ֻ��Ҫ���롣��Ҫ��һ��Ԥ��ȡ��
	// ����������ǧ�����ݡ������㻹�Ƿ���д��ɣ����ױ��ڴ�
	bool SetRangeCellString(DWORD iRowStart, DWORD iRowEnd, DWORD iColStart, DWORD iColEnd,
		const std::vector<CString>& new_string);
	// ���������ѳ���д�뷶Χ������ա������������д�뷶Χ֮ǰ�����ݡ��÷��μ�SetRangeCellString��������ȥ
	// ���磬��A2д��D10��������յ�10���Ժ�����ݡ���Ҫ��д���߲����漰ȫ��д�룬���֪����������������������ó��ġ�
	// �����10�м��Ժ�Ϊ�գ���ʲô������������
	bool SetRangeCellStringAndClearRestRows(DWORD iRowStart, DWORD iRowEnd, DWORD iColStart, DWORD iColEnd,
		const std::vector<CString>& new_string);
	// ���������ѳ���д�뷶Χ������ա������������д�뷶Χ֮ǰ�����ݡ��÷��μ�SetRangeCellString��������ȥ
	// ���磬��A2д��D10��������յ�4���Ժ�����ݡ���Ҫ��д���߲����漰ȫ��д�룬���֪����������������������ó��ġ�
	// �����4�м��Ժ�Ϊ�գ���ʲô������������
	bool SetRangeCellStringAndClearRestCols(DWORD iRowStart, DWORD iRowEnd, DWORD iColStart, DWORD iColEnd,
		const std::vector<CString>& new_string);
	// ɾ�����У����Һ������ǰ�ơ�����ע��ֻ�б���Ż���Ч��
	bool DeleteRangedCol(long iColStart, long iColEnd);
	// ɾ�����У����Һ���������ơ�����ע��ֻ�б���Ż���Ч��
	bool DeleteRangedRow(long iRowStart, long iRowEnd);
	//�õ�A1��������
	long GetA1LastRowCount();
public:
	// ��ʼ��EXCEL OLE�����һ����ڴ浱������һ��Excel��������ע��һ������ֻ����������һ��Excel App
	// ��Ҳ��Ϊʲô��Ϊ��д�˸�����ģʽ�����ࡣ
	static BOOL InitExcel();
	// �ͷ�EXCEL�� OLE���ر�Excel��������
	// �ͷ���֮���´��κ�MFCExcelFile������Ҫʹ��Excel����ҪInit��
	static void ReleaseExcel();
	// ȡ�õ�Ԫ�����ƣ����� (1,1)��ӦA1
	// ������벻�Ϸ�����ֵ������0��0����ֱ�ӷ��ؿ��ַ�����
	static CString GetCellName(long iRow, long iCol);
	// ȡ���е����ƣ�����27->AA
	static TCHAR *GetColumnName(long iColumn);
};
