#pragma once
#include <map>
#include "atlbase.h"

namespace NetOffice_ShimLoader
{
	class PathParser
	{
	public:

		PathParser();
		virtual ~PathParser();

		HRESULT Parse(BSTR path, WCHAR* result, int maxLen, BSTR documentPath);

		HRESULT ParseEx(BSTR path, BSTR subFolderPath, WCHAR* result, int maxLen, BSTR documentPath);

		HRESULT ParseEx(BSTR path, BSTR subFolderPath, BSTR filePath, WCHAR* result, int maxLen);

		HRESULT ParseEx(BSTR path, BSTR subFolderPath, BSTR filePath, WCHAR* result, int maxLen, BSTR documentPath);

	protected:

		BOOL OperatingSystemIsVistaOrAbove();

		HRESULT ParseInternal(BSTR path, WCHAR* result, int maxLen, BSTR documentPath);

		HRESULT ParseLegacyInternal(BSTR path, WCHAR* result, int maxLen, BSTR documentPath);

		GUID FindGuid(BSTR path);

		DWORD FindDWord(BSTR path);

	private:

		std::map<LPWSTR, GUID>		_parseMap;
		std::map<LPWSTR, DWORD>		_parseLegacyMap;
	};
}
