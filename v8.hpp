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

#ifndef V8_HPP
#define V8_HPP

#define UNICODE

#include <windows.h>
#include <stdio.h>

class V8
{
	private:

		IDispatch *piV8, *piCon;

		int V8::autoWrap(IDispatch* piIDispatch, int autoType, LPOLESTR name, VARIANT* pvRes, int cArgs...);

	public:

		V8::V8();
		V8::~V8();

		int V8::Initialize(wchar_t *progId);
		int V8::Connect(wchar_t *connectionString);
		int V8::EditUser(wchar_t *user, int enable1cauth, wchar_t *pass, int lockpassword, int setvisible, int enabledomainauth, wchar_t *domainaccount, int warnunsafe);
		int V8::CancelTask(wchar_t *name, int log);
		int V8::Execute(wchar_t *code);
};

#endif