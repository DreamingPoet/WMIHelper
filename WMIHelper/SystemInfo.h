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
	int init_result; // 初始化结果
private:
	IWbemServices* pSvc; //接收指向绑定到指定名称空间的IWbemServices对象的指针
	IWbemLocator* pLoc; //接收CoCreateInstance 创建的实例指针
	bool ExecQuery(const char* strQueryLanguage, const char* strQuery, IEnumWbemClassObject** pEnumerator);

};