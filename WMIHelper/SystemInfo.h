#pragma once

#include <string>
#include <tchar.h>
#include <wbemcli.h>


class SystemInfo
{
public:
	~SystemInfo();
	SystemInfo(wchar_t* WMI_Namespace = (_TCHAR*)L"ROOT\\CIMV2");
	int Get_CpuId(std::string& dest);
	int Get_DiskDriveId(std::string& dest);
	int Get_MacAddress(std::string& dest);
	int Get_OsName(std::string& dest);
	int init_result; // ��ʼ�����
private:
	IWbemServices* pSvc; //����ָ��󶨵�ָ�����ƿռ��IWbemServices�����ָ��
	IWbemLocator* pLoc; //����CoCreateInstance ������ʵ��ָ��
	bool ExecQuery(const char* strQueryLanguage, const char* strQuery, IEnumWbemClassObject** pEnumerator);

};