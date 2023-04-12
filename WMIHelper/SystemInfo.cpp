
#include <iostream>
#include "SystemInfo.h"
#include <Windows.h>
#include <comdef.h>
#include <vector>
#pragma comment(lib, "wbemuuid.lib")

using namespace std;

SystemInfo::~SystemInfo()
{
	pSvc->Release();
	pLoc->Release();
	CoUninitialize();
}

SystemInfo::SystemInfo(wchar_t* WMI_Namespace)
{
	init_result = 0;
	pSvc = NULL;
	pLoc = NULL;
	HRESULT hres;

	// Step 1: --------------------------------------------------
	// Initialize COM. ------------------------------------------
	hres = CoInitializeEx(0, COINIT_MULTITHREADED);
	if (FAILED(hres))
	{
		init_result = 1;
	}

	// Step 2: --------------------------------------------------
	// Set general COM security levels --------------------------

	hres = CoInitializeSecurity(NULL,
		-1,                          // COM authentication
		NULL,                        // Authentication services
		NULL,                        // Reserved
		RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
		RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
		NULL,                        // Authentication info
		EOAC_NONE,                   // Additional capabilities 
		NULL);                       // Reserved


	if (FAILED(hres))
	{
		cout << "Failed to initialize security. Error code = 0x"
			<< hex << hres << endl;
		CoUninitialize();
		init_result = 2;                 // Program has failed.
	}

	// Step 3: ---------------------------------------------------
	// Obtain the initial locator to WMI -------------------------

	hres = CoCreateInstance(
		CLSID_WbemLocator,
		0,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID*)&pLoc);

	if (FAILED(hres))
	{
		cout << "Failed to create IWbemLocator object."
			<< " Err code = 0x"
			<< hex << hres << endl;
		CoUninitialize();
		init_result = 3;
	}

	// Step 4: -----------------------------------------------------
	// Connect to WMI through the IWbemLocator::ConnectServer method

	//使用创建好的实例进行连接, 连接成功后的指针保存在psvc
	hres = pLoc->ConnectServer(
		_bstr_t(WMI_Namespace), // Object path of WMI namespace
		NULL,                    // User name. NULL = current user
		NULL,                    // User password. NULL = current
		0,                       // Locale. NULL indicates current
		NULL,                    // Security flags.
		0,                       // Authority (for example, Kerberos)
		0,                       // Context object 
		&pSvc                    // pointer to IWbemServices proxy
	);

	if (FAILED(hres))
	{
		cout << "Could not connect. Error code = 0x"
			<< hex << hres << endl;
		pLoc->Release();
		CoUninitialize();
		init_result = 4;
	}


	// Step 5: --------------------------------------------------
	// Set security levels on the proxy -------------------------

	hres = CoSetProxyBlanket(
		pSvc,                        // Indicates the proxy to set
		RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
		RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
		NULL,                        // Server principal name 
		RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
		RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
		NULL,                        // client identity
		EOAC_NONE                    // proxy capabilities 
	);

	if (FAILED(hres))
	{
		cout << "Could not set proxy blanket. Error code = 0x"
			<< hex << hres << endl;
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		init_result = 5;
	}

}

bool SystemInfo::ExecQuery(const char* strQueryLanguage, const char* strQuery, IEnumWbemClassObject** pEnumerator)
{

	HRESULT hres = pSvc->ExecQuery(
		bstr_t(strQueryLanguage),
		bstr_t(strQuery),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL,
		pEnumerator);

	if (FAILED(hres))
	{
		cout << "Query for operating system name failed."
			<< " Error code = 0x"
			<< hex << hres << endl;
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		init_result = 6;
		return false;               // Program has failed.
	}

	return true;
}

int SystemInfo::Get_CpuId(string& dest)
{
	IEnumWbemClassObject* pEnumerator = NULL;//存放查询语句执行结果
	ExecQuery("WQL", "SELECT * FROM Win32_Processor WHERE (ProcessorId IS NOT NULL)", &pEnumerator);

	IWbemClassObject* pclsObj = NULL;
	ULONG uReturn = 0;


	while (pEnumerator)
	{
		//  此方法将查询返回的数据对象传递给IWbemClassObject指针。
		HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
			&pclsObj, &uReturn);

		if (0 == uReturn)
		{
			break;
		}

		VARIANT vtProp;

		// Get the value of the Name property
		hr = pclsObj->Get(L"ProcessorId", 0, &vtProp, 0, 0);
		dest = (_bstr_t)vtProp.bstrVal;

		VariantClear(&vtProp);

		pclsObj->Release();
	}

	pEnumerator->Release();
	return init_result;
}

int SystemInfo::Get_DiskDriveId(string& dest)
{
	IEnumWbemClassObject* pEnumerator = NULL;//存放查询语句执行结果
	ExecQuery("WQL", "SELECT * FROM Win32_DiskDrive WHERE (SerialNumber IS NOT NULL)", &pEnumerator);

	IWbemClassObject* pclsObj = NULL;
	ULONG uReturn = 0;

	while (pEnumerator)
	{
		//  此方法将查询返回的数据对象传递给IWbemClassObject指针。
		HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
			&pclsObj, &uReturn);

		if (0 == uReturn)
		{
			break;
		}

		VARIANT vtProp;

		// Get the value of the Name property
		hr = pclsObj->Get(L"SerialNumber", 0, &vtProp, 0, 0);
		dest = (_bstr_t)vtProp.bstrVal;

		VariantClear(&vtProp);

		pclsObj->Release();

		//这里直接跳出, 只获取第一个硬盘的序列号
		break;
	}

	pEnumerator->Release();
	return init_result;
}

int SystemInfo::Get_MacAddress(string& dest)
{
	IEnumWbemClassObject* pEnumerator = NULL;//存放查询语句执行结果
	ExecQuery("WQL", "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True", &pEnumerator);

	IWbemClassObject* pclsObj = NULL;
	ULONG uReturn = 0;
	vector<string> MACAddress;

	while (pEnumerator)
	{
		//  此方法将查询返回的数据对象传递给IWbemClassObject指针。
		HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
			&pclsObj, &uReturn);

		if (0 == uReturn)
		{
			break;
		}

		VARIANT vtProp;
		pclsObj->Get(L"MACAddress", 0, &vtProp, 0, 0);
		MACAddress.push_back((string)_bstr_t(vtProp.bstrVal));

		VariantClear(&vtProp);

		pclsObj->Release();
	}

	if (MACAddress.size() >= 1)
	{
		dest = MACAddress[0];//返回第一个网卡地址
	}

	pEnumerator->Release();
	return init_result;
}

int SystemInfo::Get_OsName(string& dest)
{
	IEnumWbemClassObject* pEnumerator = NULL;//存放查询语句执行结果
	ExecQuery("WQL", "SELECT * FROM Win32_OperatingSystem WHERE (Name IS NOT NULL)", &pEnumerator);

	IWbemClassObject* pclsObj = NULL;
	ULONG uReturn = 0;

	while (pEnumerator)
	{
		//  此方法将查询返回的数据对象传递给IWbemClassObject指针。
		HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
			&pclsObj, &uReturn);

		if (0 == uReturn)
		{
			break;
		}

		VARIANT vtProp;

		// Get the value of the Name property
		hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
		dest = (_bstr_t)vtProp.bstrVal;

		VariantClear(&vtProp);

		pclsObj->Release();

		//这里直接跳出, 只获取第一个硬盘的序列号
		break;
	}

	pEnumerator->Release();
	return init_result;
}