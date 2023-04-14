//NVD 1C8 automation library
//Copyright © 2019, Nikolay Dudkin

//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:
//
//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.
//
//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.

#include "v8.hpp"

int V8::autoWrap(IDispatch *piIDispatch, int autoType, LPOLESTR name, VARIANT *pvRes, int cArgs...)
{
	va_list marker;
	va_start(marker, cArgs);

	DISPPARAMS dispParams = { NULL, NULL, 0, 0 };
	DISPID dispIdNamed = DISPID_PROPERTYPUT;
	DISPID dispId;
	VARIANT* pArgs;

	if (!piIDispatch)
		return 1;

	if (FAILED(piIDispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_SYSTEM_DEFAULT, &dispId)))
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

	if (FAILED(piIDispatch->Invoke(dispId, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dispParams, pvRes, NULL, NULL)))
		return 3;

	delete[] pArgs;

	va_end(marker);

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

	if (FAILED(CLSIDFromProgID(progId, &clsid)))
		return 110;

	if (FAILED(CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (void**)&piV8)))
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
	res = autoWrap(piV8, DISPATCH_METHOD, L"Connect", &vRes, 1);

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
	res = autoWrap(piCon, DISPATCH_PROPERTYGET, L"ПользователиИнформационнойБазы", &vRes, 0);
	if (res)
		return 300 + res;

	pUsers = vRes.pdispVal;

	VariantInit(&vStr);
	vStr.vt = VT_BSTR;
	vStr.bstrVal = ::SysAllocStringLen(user, wcslen(user));

	VariantInit(&vRes);
	res = autoWrap(pUsers, DISPATCH_METHOD, L"НайтиПоИмени", &vRes, 1);

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

		res = autoWrap(pUser, DISPATCH_PROPERTYPUT, L"АутентификацияСтандартная", NULL, 1);
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

		res = autoWrap(pUser, DISPATCH_PROPERTYPUT, L"Пароль", NULL, 1);

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

		res = autoWrap(pUser, DISPATCH_PROPERTYPUT, L"ЗапрещеноИзменятьПароль", NULL, 1);
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

		res = autoWrap(pUser, DISPATCH_PROPERTYPUT, L"ПоказыватьВСпискеВыбора", NULL, 1);
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

		res = autoWrap(pUser, DISPATCH_PROPERTYPUT, L"АутентификацияОС", NULL, 1);
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

		res = autoWrap(pUser, DISPATCH_PROPERTYPUT, L"ПользовательОС", NULL, 1);

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
		res = autoWrap(pUser, DISPATCH_PROPERTYGET, L"ЗащитаОтОпасныхДействий", &vRes, 0);
		if (res)
		{
			if (pUser)
				pUser->Release();
			return 380 + res;
		}
		
		IDispatch *pProt = vRes.pdispVal;

		res = autoWrap(pProt, DISPATCH_PROPERTYPUT, L"ПредупреждатьОбОпасныхДействиях", NULL, 1);

		if (pProt)
			pProt->Release();

		if (res)
		{
			if (pUser)
				pUser->Release();
			return 390 + res;
		}
	}

	res = autoWrap(pUser, DISPATCH_METHOD, L"Записать", NULL, 0);
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

	res = autoWrap(piCon, DISPATCH_METHOD, L"NewObject", &vStruct, 1);
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

	res = autoWrap(pStruct, DISPATCH_METHOD, L"Вставить", NULL, 2); // value - key!

	VariantClear(&vStr2);
	VariantClear(&vStr1);

	if (res)
	{
		if(pStruct)
			pStruct->Release();
		return 510 + res;
	}
	
	VariantInit(&vRes);
	res = autoWrap(piCon, DISPATCH_PROPERTYGET, L"СостояниеФоновогоЗадания", &vRes, 0);
	if (res)
	{
		if(pStruct)
			pStruct->Release();
		return 520 + res;
	}
	
	pTaskStates = vRes.pdispVal;

	VariantInit(&vTaskState);
	res = autoWrap(pTaskStates, DISPATCH_PROPERTYGET, L"Активно", &vTaskState, 0);

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

	res = autoWrap(pStruct, DISPATCH_METHOD, L"Вставить", NULL, 2); // value - key!

	VariantClear(&vStr1);
	VariantClear(&vTaskState);

	if (res)
	{
		if(pStruct)
			pStruct->Release();
		return 540 + res;
	}

	VariantInit(&vRes);
	res = autoWrap(piCon, DISPATCH_PROPERTYGET, L"ФоновыеЗадания", &vRes, 0);
	if (res)
	{
		if(pStruct)
			pStruct->Release();
		return 550 + res;
	}

	pBackTasks = vRes.pdispVal;
	
	VariantInit(&vRes);
	res = autoWrap(pBackTasks, DISPATCH_METHOD, L"ПолучитьФоновыеЗадания", &vRes, 1);

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
	res = autoWrap(pTasks, DISPATCH_METHOD, L"Количество", &vCount, 0);
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
		res = autoWrap(pTasks, DISPATCH_METHOD, L"Получить", &vRes, 1);
		if (!res)
		{
			pTask = vRes.pdispVal;
			autoWrap(pTask, DISPATCH_METHOD, L"Отменить", NULL, 0);
		
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

	int res = autoWrap(piCon, DISPATCH_METHOD, L"Exec1C", NULL, 1);

	VariantClear(&vStr);

	return res + (res > 0 ? 410 : 0);
}
