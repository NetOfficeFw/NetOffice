#include "stdafx.h"
#include "CustomRegisterValue.h"
#include <ctime>

namespace NetOffice_ShimLoader_Register
{
	/***************************************************************************
	* Ctor, Dtor
	***************************************************************************/

	CustomRegisterValue::CustomRegisterValue()
	{
		/*_name = nullptr;
		_kind = 0;
		_value = nullptr;*/
		_parseValue = FALSE;
	}

	CustomRegisterValue::CustomRegisterValue(_bstr_t name, _bstr_t kind, _bstr_t value, BOOL parseValue)
	{
		_name = name;
		_kind = kind;
		_value = value;
		_parseValue = parseValue;
	}

	CustomRegisterValue::~CustomRegisterValue()
	{
		//delete _name;
		//delete _value;
	}


	/***************************************************************************
	* CustomRegisterValue Methods
	***************************************************************************/

	BOOL CustomRegisterValue::SeemsToBeValid()
	{
		return _name.length() > 0 && _kind.length() > 0 && _value.length() > 0;
	}

	_bstr_t CustomRegisterValue::Name()
	{
		return _name;
	}

	_bstr_t CustomRegisterValue::Kind()
	{
		return _kind;
	}

	_bstr_t CustomRegisterValue::Value()
	{
		return _value;
	}

	void CustomRegisterValue::ProcessedValue(WCHAR* buffer, int maxLen)
	{
		time_t t = time(NULL);
		struct tm buf;

		if (_parseValue && 0 == wcscmp(_value, L"$LocalTime"))
		{
			if(NULL == localtime_s(&buf, &t))
				wcsftime(buffer, maxLen, L"%d-%m-%Y %I:%M:%S %p", &buf);
			else
				StringCchCopy(buffer, maxLen, _value);
		}
		else if (_parseValue && 0 == wcscmp(_value, L"$LocalTimeUSFormat"))
		{
			if (NULL == localtime_s(&buf, &t))
				wcsftime(buffer, maxLen, L"%m/%d/%Y %I:%M:%S", &buf);
			else
				StringCchCopy(buffer, maxLen, _value);
		}
		else if (_parseValue && 0 == wcscmp(_value, L"$LocalTimeDEFormat"))
		{
			if (NULL == localtime_s(&buf, &t))
				wcsftime(buffer, maxLen, L"%d.%m.%Y %H:%M:%S", &buf);
			else
				StringCchCopy(buffer, maxLen, _value);
		}
		else
		{
			StringCchCopy(buffer, maxLen, _value);
		}
	}

	DWORD CustomRegisterValue::RegKind()
	{
		if (0 == wcscmp(_kind, L"REG_DWORD"))
			return REG_DWORD;
		else
			return REG_SZ;
	}
}
