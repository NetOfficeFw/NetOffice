#pragma once
#include "stdafx.h"
#include "DllRegisterMode.hpp"
#include <map>

using namespace std;

namespace NetOffice_ShimLoader
{
	class DllRegisterModeParser
	{

	public:

		DllRegisterModeParser();
		virtual ~DllRegisterModeParser();

		RegisterMode Parse(_bstr_t text);

	private:

		map<_bstr_t, RegisterMode>	_enumMap;
		RegisterMode				_default;
	};
}
