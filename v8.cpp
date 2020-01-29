//NVD 1C8 automation library
//Copyright © 2019, Nikolay Dudkin

//This program is free software: you can redistribute it and/or modify
//it under the terms of the GNU Lesser General Public License as published by
//the Free Software Foundation, either version 3 of the License, or
//(at your option) any later version.
//This program is distributed in the hope that it will be useful,
//but WITHOUT ANY WARRANTY; without even the implied warranty of
//MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the
//GNU Lesser General Public License for more details.
//You should have received a copy of the GNU Lesser General Public License
//along with this program.If not, see<https://www.gnu.org/licenses/>.

#include "v8.hpp"

int V8::autoWrap(int autoType, VARIANT *pvRes, IDispatch *piIDispatch, LPOLESTR name, int cArgs...)
{
	va_list marker;
	va_start(marker, cArgs);

	if (!piIDispatch)
		return 1;

	DISPPARAMS dispParams = {NULL, NULL, 0, 0};
	DISPID dispIdNamed = DISPID_PROPERTYPUT;
	DISPID dispId;

	HRESULT hr;
	VARIANT *pArgs;

	hr = piIDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_SYSTEM_DEFAULT, &dispId);
	if (FAILED(hr))
		return 2;

	pArgs = new VARIANT[cArgs + 1];

	for (int i = 0; i < cArgs; i++)
	{
		pArgs[i] = va_arg(marker, VARIANT);
	}

	dispParams.cArgs = cArgs;
	dispParams.rgvarg = pArgs;

	if (autoType & DISPATCH_PROPERTYPUT)
	{
		dispParams.cNamedArgs = 1;
		dispParams.rgdispidNamedArgs = &dispIdNamed;
	}

	hr = piIDispatch->Invoke(dispId, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dispParams, pvRes, NULL, NULL);
	if (FAILED(hr))
		return 3;

	va_end(marker);

	delete[] pArgs;

	return 0;
}

V8::V8()
{
	piV8 = NULL;
	piCon = NULL;
}

V8::~V8()
{
	if (piCon)
		piCon->Release();
	if (piV8)
		piV8->Release();
}

int V8::Initialize(wchar_t *progId)
{
	CLSID clsid;

	HRESULT hr = CLSIDFromProgID(progId, &clsid);
	if (FAILED(hr))
		return 110;

	hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (void **)&piV8);
	if (FAILED(hr))
	{
		return 120;
	}

	return 0;
}

int V8::Connect(wchar_t *connectionString)
{
	int res = 0;

	VARIANT vStr;
	VARIANT vRes;
	
	VariantInit(&vStr);
	vStr.vt = VT_BSTR;
	vStr.bstrVal = ::SysAllocStringLen(connectionString, wcslen(connectionString));

	VariantInit(&vRes);
	res = autoWrap(DISPATCH_METHOD, &vRes, piV8, L"Connect", 1, vStr);

	VariantClear(&vStr);

	if (res)
		return 200 + res;

	piCon = vRes.pdispVal;

	return res;
}

int V8::EditUser(wchar_t *user, int enable1cauth, wchar_t *pass, int lockpassword, int setvisible, int enabledomainauth, wchar_t *domainaccount, int warnunsafe)
{
	int res = 0;

	VARIANT vStr;
	VARIANT vBool;
	VARIANT vRes;
	
	IDispatch *pUsers = NULL;
	IDispatch *pUser = NULL;

	VariantInit(&vRes);
	res = autoWrap(DISPATCH_PROPERTYGET, &vRes, piCon, L"ПользователиИнформационнойБазы", 0);
	if (res)
		return 300 + res;

	pUsers = vRes.pdispVal;

	VariantInit(&vStr);
	vStr.vt = VT_BSTR;
	vStr.bstrVal = ::SysAllocStringLen(user, wcslen(user));

	VariantInit(&vRes);
	res = autoWrap(DISPATCH_METHOD, &vRes, pUsers, L"НайтиПоИмени", 1, vStr);

	VariantClear(&vStr);

	if (res)
	{
		if (pUsers)
			pUsers->Release();
		return 310 + res;
	}

	pUser = vRes.pdispVal;
	
	if (pUsers)
		pUsers->Release();

	VariantInit(&vBool);
	vBool.vt = VT_BOOL;

	if(enable1cauth >= 0)
	{
		vBool.boolVal = enable1cauth;

		res = autoWrap(DISPATCH_PROPERTYPUT, NULL, pUser, L"АутентификацияСтандартная", 1, vBool);
		if (res)
		{
			if (pUser)
				pUser->Release();
			return 320 + res;
		}
	}

	if(pass)
	{
		VariantInit(&vStr);
		vStr.vt = VT_BSTR;
		vStr.bstrVal = ::SysAllocStringLen(pass, wcslen(pass));

		res = autoWrap(DISPATCH_PROPERTYPUT, NULL, pUser, L"Пароль", 1, vStr);

		VariantClear(&vStr);

		if (res)
		{
			if (pUser)
				pUser->Release();
			return 330 + res;
		}
	}

	if(lockpassword >= 0)
	{
		vBool.boolVal = lockpassword;

		res = autoWrap(DISPATCH_PROPERTYPUT, NULL, pUser, L"ЗапрещеноИзменятьПароль", 1, vBool);
		if (res)
		{
			if (pUser)
				pUser->Release();
			return 340 + res;
		}
	}

	if(setvisible >= 0)
	{
		vBool.boolVal = setvisible;

		res = autoWrap(DISPATCH_PROPERTYPUT, NULL, pUser, L"ПоказыватьВСпискеВыбора", 1, vBool);
		if (res)
		{
			if (pUser)
				pUser->Release();
			return 350 + res;
		}
	}

	if(enabledomainauth >= 0)
	{
		vBool.boolVal = enabledomainauth;

		res = autoWrap(DISPATCH_PROPERTYPUT, NULL, pUser, L"АутентификацияОС", 1, vBool);
		if (res)
		{
			if (pUser)
				pUser->Release();
			return 360 + res;
		}
	}

	if(domainaccount)
	{
		VariantInit(&vStr);
		vStr.vt = VT_BSTR;
		vStr.bstrVal = ::SysAllocStringLen(domainaccount, wcslen(domainaccount));

		res = autoWrap(DISPATCH_PROPERTYPUT, NULL, pUser, L"ПользовательОС", 1, vStr);

		VariantClear(&vStr);

		if (res)
		{
			if (pUser)
				pUser->Release();
			return 370 + res;
		}
	}

	if(warnunsafe >= 0)
	{
		vBool.boolVal = warnunsafe;

		VariantInit(&vRes);
		res = autoWrap(DISPATCH_PROPERTYGET, &vRes, pUser, L"ЗащитаОтОпасныхДействий", 0);
		if (res)
		{
			if (pUser)
				pUser->Release();
			return 380 + res;
		}
		
		IDispatch *pProt = vRes.pdispVal;

		res = autoWrap(DISPATCH_PROPERTYPUT, NULL, pProt, L"ПредупреждатьОбОпасныхДействиях", 1, vBool);

		if (pProt)
			pProt->Release();

		if (res)
		{
			if (pUser)
				pUser->Release();
			return 390 + res;
		}
	}

	res = autoWrap(DISPATCH_METHOD, NULL, pUser, L"Записать", 0);
	if (res)
	{
		if (pUser)
			pUser->Release();
		return 400 + res;
	}

	if (pUser)
		pUser->Release();

	return 0;
}

int V8::CancelTask(wchar_t *name, int log)
{
	int res = 0;
	
	VARIANT vStr1;
	VARIANT vStr2;
	VARIANT vStruct;
	VARIANT vTaskState;
	VARIANT vCount;
	VARIANT vRes;
	
	IDispatch *pStruct = NULL;
	IDispatch *pTaskStates = NULL;
	IDispatch *pBackTasks = NULL;
	IDispatch *pTasks = NULL;
	IDispatch *pTask = NULL;
	
	VariantInit(&vStr1);
	vStr1.vt = VT_BSTR;
	vStr1.bstrVal = ::SysAllocStringLen(L"Структура", wcslen(L"Структура"));
	
	VariantInit(&vStruct);

	res = autoWrap(DISPATCH_METHOD, &vStruct, piCon, L"NewObject", 1, vStr1);
	if (res)
		return 500 + res;

	VariantClear(&vStr1);

	pStruct = vStruct.pdispVal;

	VariantInit(&vStr1);
	vStr1.vt = VT_BSTR;
	vStr1.bstrVal = ::SysAllocStringLen(name, wcslen(name));

	VariantInit(&vStr2);
	vStr2.vt = VT_BSTR;
	vStr2.bstrVal = ::SysAllocStringLen(L"Наименование", wcslen(L"Наименование"));

	res = autoWrap(DISPATCH_METHOD, NULL, pStruct, L"Вставить", 2, vStr1, vStr2); // value - key!

	VariantClear(&vStr2);
	VariantClear(&vStr1);

	if (res)
	{
		if(pStruct)
			pStruct->Release();
		return 510 + res;
	}
	
	VariantInit(&vRes);
	res = autoWrap(DISPATCH_PROPERTYGET, &vRes, piCon, L"СостояниеФоновогоЗадания", 0);
	if (res)
	{
		if(pStruct)
			pStruct->Release();
		return 520 + res;
	}
	
	pTaskStates = vRes.pdispVal;

	VariantInit(&vTaskState);
	res = autoWrap(DISPATCH_PROPERTYGET, &vTaskState, pTaskStates, L"Активно", 0);

	if(pTaskStates)
		pTaskStates->Release();

	if (res)
	{
		if(pStruct)
			pStruct->Release();
		return 530 + res;
	}
	
	VariantInit(&vStr1);
	vStr1.vt = VT_BSTR;
	vStr1.bstrVal = ::SysAllocStringLen(L"Состояние", wcslen(L"Состояние"));

	res = autoWrap(DISPATCH_METHOD, NULL, pStruct, L"Вставить", 2, vTaskState, vStr1); // value - key!

	VariantClear(&vStr1);
	VariantClear(&vTaskState);

	if (res)
	{
		if(pStruct)
			pStruct->Release();
		return 540 + res;
	}

	VariantInit(&vRes);
	res = autoWrap(DISPATCH_PROPERTYGET, &vRes, piCon, L"ФоновыеЗадания", 0);
	if (res)
	{
		if(pStruct)
			pStruct->Release();
		return 550 + res;
	}

	pBackTasks = vRes.pdispVal;
	
	VariantInit(&vRes);
	res = autoWrap(DISPATCH_METHOD, &vRes, pBackTasks, L"ПолучитьФоновыеЗадания", 1, vStruct);

	if(pBackTasks)
		pBackTasks->Release();

	if(pStruct)
		pStruct->Release();

	if (res)
	{
		return 560 + res;
	}
	
	pTasks = vRes.pdispVal;
	
	VariantInit(&vCount);
	res = autoWrap(DISPATCH_METHOD, &vCount, pTasks, L"Количество", 0);
	if (res)
	{
		if(pTasks)
			pTasks->Release();
		return 570 + res;
	}
	
	int count = vCount.iVal;
	
	if(log)
		fwprintf(stderr, L"victims: %d\r\n", count);
	
	for(int i = 0; i < count; i++)
	{
		vCount.iVal = i;
		
		VariantInit(&vRes);
		res = autoWrap(DISPATCH_METHOD, &vRes, pTasks, L"Получить", 1, vCount);
		if (!res)
		{
			pTask = vRes.pdispVal;
			autoWrap(DISPATCH_METHOD, NULL, pTask, L"Отменить", 0);
		
			if(pTask)
				pTask->Release();
		}		
	}

	if(pTasks)
		pTasks->Release();

	return 0;
}

int V8::Execute(wchar_t *code)
{
	wchar_t *buf = ::SysAllocStringLen(code, wcslen(code));

	VARIANT vStr;
	vStr.vt = VT_BSTR;
	vStr.bstrVal = buf;

	int res = autoWrap(DISPATCH_METHOD, NULL, piCon, L"Exec1C", 1, vStr);

	VariantClear(&vStr);

	return res + (res > 0 ? 410 : 0);
}
