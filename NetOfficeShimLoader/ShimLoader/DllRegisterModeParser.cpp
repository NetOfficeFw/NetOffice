#include "stdafx.h"
#include "DllRegisterModeParser.h"

namespace NetOffice_ShimLoader
{
	DllRegisterModeParser::DllRegisterModeParser()
	{
		_enumMap["SystemComponentAndUserAddin"] = SystemComponentAndUserAddin;
		_enumMap["System"] = System;
		_enumMap["User"] = User;
		_default = SystemComponentAndUserAddin;
	}

	DllRegisterModeParser::~DllRegisterModeParser()
	{
		_enumMap.clear();
	}

	RegisterMode DllRegisterModeParser::Parse(_bstr_t text)
	{
		if (_enumMap.find(text) == _enumMap.end())
			return  _enumMap.find(text)->second;
		else
			return _default;
	}
}
