#include "stdafx.h"
#include "CustomRegisterValue.h"

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
	}

	CustomRegisterValue::CustomRegisterValue(_bstr_t name, _bstr_t kind, _bstr_t value)
	{
		_name = name;
		_kind = kind;
		_value = value;
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

	DWORD CustomRegisterValue::RegKind()
	{
		if (0 == wcscmp(_kind, L"REG_DWORD"))
			return REG_DWORD;
		else
			return REG_SZ;
	}
}
