#pragma once

namespace NetOffice_ShimLoader_Register
{
	class CustomRegisterValue
	{
	public:

		// Ctor, Dtor
		CustomRegisterValue();
		CustomRegisterValue(_bstr_t name, _bstr_t kind, _bstr_t value, BOOL parseValue);
		virtual ~CustomRegisterValue();

		// CustomRegisterValue Methods
		BOOL SeemsToBeValid();
		_bstr_t Name();
		_bstr_t Kind();
		_bstr_t Value();
		void ProcessedValue(WCHAR* buffer, int maxLen);

		DWORD RegKind();

	private:

		_bstr_t _name;
		_bstr_t _kind;
		_bstr_t _value;
		BOOL	_parseValue;
	};

	// Pointer to CustomRegisterValue
	typedef CustomRegisterValue* PCustomRegisterValue;
}
