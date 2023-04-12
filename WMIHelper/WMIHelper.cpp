
#define _WIN32_DCOM
#include <iostream>
using namespace std;
#include <comdef.h>
#include <Wbemidl.h>
#include <string>

#pragma comment(lib, "wbemuuid.lib")

int InitCOM()
{

	HRESULT hres;
	// Step 1: --------------------------------------------------
	// Initialize COM. ------------------------------------------

	hres = CoInitializeEx(0, COINIT_MULTITHREADED);
	if (FAILED(hres))
	{		
		return 1; // Program has failed.
	}

	// Step 2: --------------------------------------------------
	// Set general COM security levels --------------------------

	hres = CoInitializeSecurity(
		NULL,
		-1,                          // COM authentication
		NULL,                        // Authentication services
		NULL,                        // Reserved
		RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
		RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
		NULL,                        // Authentication info
		EOAC_NONE,                   // Additional capabilities 
		NULL                         // Reserved
	);


	if (FAILED(hres))
	{
		CoUninitialize();
		return 2;   // Program has failed.
	}
	return 0;
}
	
int WMIQuery(wchar_t* pchar, const char* strQuery, const wchar_t* keyWords)
{

	HRESULT hres;

	// Step 3: ---------------------------------------------------
	// Obtain the initial locator to WMI -------------------------

	IWbemLocator* pLoc = NULL;

	hres = CoCreateInstance(
		CLSID_WbemLocator,
		0,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID*)&pLoc);

	if (FAILED(hres))
	{
		CoUninitialize();
		return 3;               // Program has failed.
	}

	// Step 4: -----------------------------------------------------
	// Connect to WMI through the IWbemLocator::ConnectServer method

	IWbemServices* pSvc = NULL;

	// Connect to the root\cimv2 namespace with
	// the current user and obtain pointer pSvc
	// to make IWbemServices calls.
	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
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

		pLoc->Release();
		CoUninitialize();

		return 4;                // Program has failed.
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
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return 5;             // Program has failed.
	}

	// Step 6: --------------------------------------------------
	// Use the IWbemServices pointer to make requests of WMI ----

	// For example, get the name of the operating system
	IEnumWbemClassObject* pEnumerator = NULL;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t(strQuery),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL,
		&pEnumerator);

	if (FAILED(hres))
	{
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return 6;             // Program has failed.
	}

	// Step 7: -------------------------------------------------
	// Get the data from the query in step 6 -------------------

	IWbemClassObject* pclsObj = NULL;
	ULONG uReturn = 0;

	while (pEnumerator)
	{
		HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
			&pclsObj, &uReturn);

		if (0 == uReturn)
		{
			break;
		}

		VARIANT vtProp;


		// Get the value of the Name property
		hr = pclsObj->Get(keyWords, 0, &vtProp, 0, 0);
		
		wstring temp;
		switch (vtProp.vt)
		{
		case VT_I2:
			temp = to_wstring(vtProp.uiVal);
			break;
		case VT_I4:
			temp = to_wstring(vtProp.ulVal);
			break;
		case VT_BSTR:
			temp = vtProp.bstrVal;
			break;
		case VT_BOOL:
			temp = to_wstring(vtProp.boolVal);
			break;
		case VT_I1:
			temp = to_wstring(vtProp.cVal);
			break;
		case VT_UI1:
			temp = to_wstring(vtProp.bVal);
			break;
		case VT_UI2:
			temp = to_wstring(vtProp.uiVal);
			break;
		case VT_UI4:
			temp = to_wstring(vtProp.ulVal);
			break;
		case VT_I8:
			temp = to_wstring(vtProp.uiVal);
			break;
		case VT_UI8:
			temp = to_wstring(vtProp.llVal);
			break;
		default:
			break;
		}

	
		wcscpy_s(pchar, 128, temp.c_str());
	
		VariantClear(&vtProp);
		pclsObj->Release();
	}


	// Cleanup
	// ========
	pSvc->Release();
	pLoc->Release();
	pEnumerator->Release();
	return 0;
}


void UnInit()
{
	CoUninitialize();  // Program successfully completed.
}


int main()
{

	InitCOM();
	wchar_t a[128]{ L"\0"};
	//int c = WMIQuery(a, "SELECT * FROM Win32_VideoController", L"AdapterRAM");
	int c = WMIQuery(a, "SELECT * FROM Win32_ComputerSystem", L"TotalPhysicalMemory");

	wcout << a << endl;

	UnInit();

	string str;
	cin >> str;
	return 0;

}

