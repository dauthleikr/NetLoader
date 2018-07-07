// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"
#include <metahost.h>
#include <tchar.h>
#include <string>
#include <thread>
#include <codecvt>
#include <comdef.h>
#include <corerror.h>

using namespace std;

#pragma comment(lib, "mscoree.lib")

ICLRMetaHost *metaHost = nullptr;
ICLRRuntimeInfo *runtimeInfo = nullptr;
ICLRRuntimeHost *runtimeHost = nullptr;

struct AssemblyInfo
{
	LPCWSTR AssemblyPath;
	LPCWSTR TypeName;
	LPCWSTR MethodName;
	LPCWSTR Param;
};

BOOL APIENTRY DllMain(HMODULE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved)
{
	if (ul_reason_for_call == DLL_PROCESS_ATTACH)
	{
		// entry code?
	}

	return TRUE;
}


HMODULE GetCurrentModule()
{
	HMODULE hModule = nullptr;
	GetModuleHandleEx(
		GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,
		reinterpret_cast<LPCTSTR>(GetCurrentModule),
		&hModule);

	return hModule;
}

BOOL PrintErrorOnBadHResult(HRESULT result, LPCWSTR resultName)
{
	switch (result)
	{
	case S_OK:
		return FALSE;
	case HOST_E_CLRNOTAVAILABLE:
		MessageBoxW(nullptr, resultName, L"HOST_E_CLRNOTAVAILABLE", 0);
		break;
	case HOST_E_TIMEOUT:
		MessageBoxW(nullptr, resultName, L"HOST_E_TIMEOUT", 0);
		break;
	case HOST_E_NOT_OWNER:
		MessageBoxW(nullptr, resultName, L"HOST_E_NOT_OWNER", 0);
		break;
	case HOST_E_ABANDONED:
		MessageBoxW(nullptr, resultName, L"HOST_E_ABANDONED", 0);
		break;
	case E_FAIL:
		MessageBoxW(nullptr, resultName, L"E_FAIL", 0);
		break;
	case 1: // Already running
		break;
	default:
		MessageBoxW(nullptr, resultName, (wstring(L"UNKNOWN ERROR: ") + to_wstring(result)).c_str(), 0);
		break;
	}
	return TRUE;
}

extern "C"
{
	__declspec(dllexport) DWORD WINAPI LoadAssembly(LPVOID assemblyInfoPtr)
	{
		auto *assemblyInfo = static_cast<AssemblyInfo*>(assemblyInfoPtr);

		//MessageBoxW(NULL, L"Loading assembly", L"sugoi", 0);
		//MessageBoxW(NULL, assemblyInfo->AssemblyPath, L"asm name", 0);
		//MessageBoxW(NULL, assemblyInfo->MethodName, L"method name", 0);
		//MessageBoxW(NULL, assemblyInfo->TypeName, L"type name", 0);

		DWORD returnValue;
		DWORD hresult;

		hresult = runtimeHost->ExecuteInDefaultAppDomain(assemblyInfo->AssemblyPath, assemblyInfo->TypeName, assemblyInfo->MethodName, assemblyInfo->Param, &returnValue);

		//wchar_t buffer[256];
		//wsprintfW(buffer, L"%d", returnValue);
		//MessageBoxW(nullptr, buffer, buffer, MB_OK);

		if (PrintErrorOnBadHResult(hresult, L"ExecuteInDefaultAppDomain")) return -1;

		return returnValue;
	}

	__declspec(dllexport) DWORD WINAPI StartCLR(LPVOID runtimeVersion)
	{
		auto instanceResult = CLRCreateInstance(CLSID_CLRMetaHost, IID_PPV_ARGS(&metaHost));
		if (PrintErrorOnBadHResult(instanceResult, L"CLRCreateInstance")) return instanceResult;

		auto runtimeResult = metaHost->GetRuntime(static_cast<LPCWSTR>(runtimeVersion), IID_PPV_ARGS(&runtimeInfo));
		if (PrintErrorOnBadHResult(runtimeResult, L"GetRuntime")) return runtimeResult;

		auto interfaceResult = runtimeInfo->GetInterface(CLSID_CLRRuntimeHost, IID_PPV_ARGS(&runtimeHost));
		if (PrintErrorOnBadHResult(interfaceResult, L"GetInterface")) return interfaceResult;

		auto startResult = runtimeHost->Start();
		if (PrintErrorOnBadHResult(startResult, L"Start")) return startResult;

		return startResult;
	}

	__declspec(dllexport) DWORD WINAPI StopCLR(LPVOID _)
	{
		if (runtimeHost == nullptr) return -1;

		auto stopResult = runtimeHost->Stop();
		if (PrintErrorOnBadHResult(stopResult, L"Stop")) return stopResult;

		return runtimeHost->Stop();
	}

	__declspec(dllexport) DWORD WINAPI UnloadSelf(LPVOID _)
	{
		auto mod = GetCurrentModule();

		FreeLibrary(mod);
		FreeLibraryAndExitThread(mod, 0);
	}
}

