using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IActiveIMMApp 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("08C0E040-62D1-11D1-9326-0060B067B86E")]
	public interface IActiveIMMApp : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="hIME">Int32 hIME</param>
		/// <param name="phPrev">Int32 phPrev</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 AssociateContext(_RemotableHandle hWnd, Int32 hIME, out Int32 phPrev);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="dwMode">Int32 dwMode</param>
		/// <param name="pData">__MIDL___MIDL_itf_mshtml_0001_0042_0001 pData</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 ConfigureIMEA(object hKL, _RemotableHandle hWnd, Int32 dwMode, __MIDL___MIDL_itf_mshtml_0001_0042_0001 pData);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="dwMode">Int32 dwMode</param>
		/// <param name="pData">__MIDL___MIDL_itf_mshtml_0001_0042_0002 pData</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 ConfigureIMEW(object hKL, _RemotableHandle hWnd, Int32 dwMode, __MIDL___MIDL_itf_mshtml_0001_0042_0002 pData);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="phIMC">Int32 phIMC</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 CreateContext(out Int32 phIMC);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIME">Int32 hIME</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 DestroyContext(Int32 hIME);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="szReading">string szReading</param>
		/// <param name="dwStyle">Int32 dwStyle</param>
		/// <param name="szRegister">string szRegister</param>
		/// <param name="pData">object pData</param>
		/// <param name="pEnum">NetOffice.MSHTMLApi.IEnumRegisterWordA pEnum</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 EnumRegisterWordA(object hKL, string szReading, Int32 dwStyle, string szRegister, object pData, out NetOffice.MSHTMLApi.IEnumRegisterWordA pEnum);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="szReading">string szReading</param>
		/// <param name="dwStyle">Int32 dwStyle</param>
		/// <param name="szRegister">string szRegister</param>
		/// <param name="pData">object pData</param>
		/// <param name="pEnum">NetOffice.MSHTMLApi.IEnumRegisterWordW pEnum</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 EnumRegisterWordW(object hKL, string szReading, Int32 dwStyle, string szRegister, object pData, out NetOffice.MSHTMLApi.IEnumRegisterWordW pEnum);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="uEscape">UIntPtr uEscape</param>
		/// <param name="pData">object pData</param>
		/// <param name="plResult">Int32 plResult</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 EscapeA(object hKL, Int32 hIMC, UIntPtr uEscape, object pData, out Int32 plResult);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="uEscape">UIntPtr uEscape</param>
		/// <param name="pData">object pData</param>
		/// <param name="plResult">Int32 plResult</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 EscapeW(object hKL, Int32 hIMC, UIntPtr uEscape, object pData, out Int32 plResult);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="pCandList">__MIDL___MIDL_itf_mshtml_0001_0042_0007 pCandList</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetCandidateListA(Int32 hIMC, Int32 dwIndex, UIntPtr uBufLen, out __MIDL___MIDL_itf_mshtml_0001_0042_0007 pCandList, out UIntPtr puCopied);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="pCandList">__MIDL___MIDL_itf_mshtml_0001_0042_0007 pCandList</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetCandidateListW(Int32 hIMC, Int32 dwIndex, UIntPtr uBufLen, out __MIDL___MIDL_itf_mshtml_0001_0042_0007 pCandList, out UIntPtr puCopied);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pdwListSize">Int32 pdwListSize</param>
		/// <param name="pdwBufLen">Int32 pdwBufLen</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetCandidateListCountA(Int32 hIMC, out Int32 pdwListSize, out Int32 pdwBufLen);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pdwListSize">Int32 pdwListSize</param>
		/// <param name="pdwBufLen">Int32 pdwBufLen</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetCandidateListCountW(Int32 hIMC, out Int32 pdwListSize, out Int32 pdwBufLen);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="pCandidate">__MIDL___MIDL_itf_mshtml_0001_0042_0005 pCandidate</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetCandidateWindow(Int32 hIMC, Int32 dwIndex, out __MIDL___MIDL_itf_mshtml_0001_0042_0005 pCandidate);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="plf">__MIDL___MIDL_itf_mshtml_0001_0042_0003 plf</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetCompositionFontA(Int32 hIMC, out __MIDL___MIDL_itf_mshtml_0001_0042_0003 plf);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="plf">__MIDL___MIDL_itf_mshtml_0001_0042_0004 plf</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetCompositionFontW(Int32 hIMC, out __MIDL___MIDL_itf_mshtml_0001_0042_0004 plf);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="dwBufLen">Int32 dwBufLen</param>
		/// <param name="plCopied">Int32 plCopied</param>
		/// <param name="pBuf">object pBuf</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetCompositionStringA(Int32 hIMC, Int32 dwIndex, Int32 dwBufLen, out Int32 plCopied, out object pBuf);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="dwBufLen">Int32 dwBufLen</param>
		/// <param name="plCopied">Int32 plCopied</param>
		/// <param name="pBuf">object pBuf</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetCompositionStringW(Int32 hIMC, Int32 dwIndex, Int32 dwBufLen, out Int32 plCopied, out object pBuf);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pCompForm">__MIDL___MIDL_itf_mshtml_0001_0042_0006 pCompForm</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetCompositionWindow(Int32 hIMC, out __MIDL___MIDL_itf_mshtml_0001_0042_0006 pCompForm);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="phIMC">Int32 phIMC</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetContext(_RemotableHandle hWnd, out Int32 phIMC);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pSrc">string pSrc</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="uFlag">UIntPtr uFlag</param>
		/// <param name="pDst">__MIDL___MIDL_itf_mshtml_0001_0042_0007 pDst</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetConversionListA(object hKL, Int32 hIMC, string pSrc, UIntPtr uBufLen, UIntPtr uFlag, out __MIDL___MIDL_itf_mshtml_0001_0042_0007 pDst, out UIntPtr puCopied);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pSrc">string pSrc</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="uFlag">UIntPtr uFlag</param>
		/// <param name="pDst">__MIDL___MIDL_itf_mshtml_0001_0042_0007 pDst</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetConversionListW(object hKL, Int32 hIMC, string pSrc, UIntPtr uBufLen, UIntPtr uFlag, out __MIDL___MIDL_itf_mshtml_0001_0042_0007 pDst, out UIntPtr puCopied);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pfdwConversion">Int32 pfdwConversion</param>
		/// <param name="pfdwSentence">Int32 pfdwSentence</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetConversionStatus(Int32 hIMC, out Int32 pfdwConversion, out Int32 pfdwSentence);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="phDefWnd">_RemotableHandle phDefWnd</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetDefaultIMEWnd(_RemotableHandle hWnd, out _RemotableHandle phDefWnd);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="szDescription">string szDescription</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetDescriptionA(object hKL, UIntPtr uBufLen, out string szDescription, out UIntPtr puCopied);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="szDescription">string szDescription</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetDescriptionW(object hKL, UIntPtr uBufLen, out string szDescription, out UIntPtr puCopied);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="dwBufLen">Int32 dwBufLen</param>
		/// <param name="pBuf">string pBuf</param>
		/// <param name="pdwResult">Int32 pdwResult</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetGuideLineA(Int32 hIMC, Int32 dwIndex, Int32 dwBufLen, out string pBuf, out Int32 pdwResult);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="dwBufLen">Int32 dwBufLen</param>
		/// <param name="pBuf">string pBuf</param>
		/// <param name="pdwResult">Int32 pdwResult</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetGuideLineW(Int32 hIMC, Int32 dwIndex, Int32 dwBufLen, out string pBuf, out Int32 pdwResult);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="szFileName">string szFileName</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetIMEFileNameA(object hKL, UIntPtr uBufLen, out string szFileName, out UIntPtr puCopied);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="szFileName">string szFileName</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetIMEFileNameW(object hKL, UIntPtr uBufLen, out string szFileName, out UIntPtr puCopied);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetOpenStatus(Int32 hIMC);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="fdwIndex">Int32 fdwIndex</param>
		/// <param name="pdwProperty">Int32 pdwProperty</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetProperty(object hKL, Int32 fdwIndex, out Int32 pdwProperty);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="nItem">UIntPtr nItem</param>
		/// <param name="pStyleBuf">__MIDL___MIDL_itf_mshtml_0001_0042_0008 pStyleBuf</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetRegisterWordStyleA(object hKL, UIntPtr nItem, out __MIDL___MIDL_itf_mshtml_0001_0042_0008 pStyleBuf, out UIntPtr puCopied);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="nItem">UIntPtr nItem</param>
		/// <param name="pStyleBuf">__MIDL___MIDL_itf_mshtml_0001_0042_0009 pStyleBuf</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetRegisterWordStyleW(object hKL, UIntPtr nItem, out __MIDL___MIDL_itf_mshtml_0001_0042_0009 pStyleBuf, out UIntPtr puCopied);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pptPos">tagPOINT pptPos</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetStatusWindowPos(Int32 hIMC, out tagPOINT pptPos);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="puVirtualKey">UIntPtr puVirtualKey</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetVirtualKey(_RemotableHandle hWnd, out UIntPtr puVirtualKey);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="szIMEFileName">string szIMEFileName</param>
		/// <param name="szLayoutText">string szLayoutText</param>
		/// <param name="phKL">object phKL</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 InstallIMEA(string szIMEFileName, string szLayoutText, out object phKL);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="szIMEFileName">string szIMEFileName</param>
		/// <param name="szLayoutText">string szLayoutText</param>
		/// <param name="phKL">object phKL</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 InstallIMEW(string szIMEFileName, string szLayoutText, out object phKL);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsIME(object hKL);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWndIME">_RemotableHandle hWndIME</param>
		/// <param name="msg">UIntPtr msg</param>
		/// <param name="wParam">Int32 wParam</param>
		/// <param name="lParam">Int32 lParam</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsUIMessageA(_RemotableHandle hWndIME, UIntPtr msg, Int32 wParam, Int32 lParam);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWndIME">_RemotableHandle hWndIME</param>
		/// <param name="msg">UIntPtr msg</param>
		/// <param name="wParam">Int32 wParam</param>
		/// <param name="lParam">Int32 lParam</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 IsUIMessageW(_RemotableHandle hWndIME, UIntPtr msg, Int32 wParam, Int32 lParam);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwAction">Int32 dwAction</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="dwValue">Int32 dwValue</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 NotifyIME(Int32 hIMC, Int32 dwAction, Int32 dwIndex, Int32 dwValue);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="szReading">string szReading</param>
		/// <param name="dwStyle">Int32 dwStyle</param>
		/// <param name="szRegister">string szRegister</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 RegisterWordA(object hKL, string szReading, Int32 dwStyle, string szRegister);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="szReading">string szReading</param>
		/// <param name="dwStyle">Int32 dwStyle</param>
		/// <param name="szRegister">string szRegister</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 RegisterWordW(object hKL, string szReading, Int32 dwStyle, string szRegister);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="hIMC">Int32 hIMC</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 ReleaseContext(_RemotableHandle hWnd, Int32 hIMC);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pCandidate">__MIDL___MIDL_itf_mshtml_0001_0042_0005 pCandidate</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetCandidateWindow(Int32 hIMC, __MIDL___MIDL_itf_mshtml_0001_0042_0005 pCandidate);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="plf">__MIDL___MIDL_itf_mshtml_0001_0042_0003 plf</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetCompositionFontA(Int32 hIMC, __MIDL___MIDL_itf_mshtml_0001_0042_0003 plf);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="plf">__MIDL___MIDL_itf_mshtml_0001_0042_0004 plf</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetCompositionFontW(Int32 hIMC, __MIDL___MIDL_itf_mshtml_0001_0042_0004 plf);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="pComp">object pComp</param>
		/// <param name="dwCompLen">Int32 dwCompLen</param>
		/// <param name="pRead">object pRead</param>
		/// <param name="dwReadLen">Int32 dwReadLen</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetCompositionStringA(Int32 hIMC, Int32 dwIndex, object pComp, Int32 dwCompLen, object pRead, Int32 dwReadLen);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="pComp">object pComp</param>
		/// <param name="dwCompLen">Int32 dwCompLen</param>
		/// <param name="pRead">object pRead</param>
		/// <param name="dwReadLen">Int32 dwReadLen</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetCompositionStringW(Int32 hIMC, Int32 dwIndex, object pComp, Int32 dwCompLen, object pRead, Int32 dwReadLen);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pCompForm">__MIDL___MIDL_itf_mshtml_0001_0042_0006 pCompForm</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetCompositionWindow(Int32 hIMC, __MIDL___MIDL_itf_mshtml_0001_0042_0006 pCompForm);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="fdwConversion">Int32 fdwConversion</param>
		/// <param name="fdwSentence">Int32 fdwSentence</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetConversionStatus(Int32 hIMC, Int32 fdwConversion, Int32 fdwSentence);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="fOpen">Int32 fOpen</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetOpenStatus(Int32 hIMC, Int32 fOpen);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pptPos">tagPOINT pptPos</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SetStatusWindowPos(Int32 hIMC, tagPOINT pptPos);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="dwHotKeyID">Int32 dwHotKeyID</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 SimulateHotKey(_RemotableHandle hWnd, Int32 dwHotKeyID);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="szReading">string szReading</param>
		/// <param name="dwStyle">Int32 dwStyle</param>
		/// <param name="szUnregister">string szUnregister</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 UnregisterWordA(object hKL, string szReading, Int32 dwStyle, string szUnregister);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="szReading">string szReading</param>
		/// <param name="dwStyle">Int32 dwStyle</param>
		/// <param name="szUnregister">string szUnregister</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 UnregisterWordW(object hKL, string szReading, Int32 dwStyle, string szUnregister);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fRestoreLayout">Int32 fRestoreLayout</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 Activate(Int32 fRestoreLayout);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		Int32 Deactivate();

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="msg">UIntPtr msg</param>
		/// <param name="wParam">Int32 wParam</param>
		/// <param name="lParam">Int32 lParam</param>
		/// <param name="plResult">Int32 plResult</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 OnDefWindowProc(_RemotableHandle hWnd, UIntPtr msg, Int32 wParam, Int32 lParam, out Int32 plResult);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="aaClassList">Int16 aaClassList</param>
		/// <param name="uSize">UIntPtr uSize</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 FilterClientWindows(Int16 aaClassList, UIntPtr uSize);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="uCodePage">UIntPtr uCodePage</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetCodePageA(object hKL, out UIntPtr uCodePage);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="plid">Int16 plid</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetLangId(object hKL, out Int16 plid);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 AssociateContextEx(_RemotableHandle hWnd, Int32 hIMC, Int32 dwFlags);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="idThread">Int32 idThread</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 DisableIME(Int32 idThread);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="dwType">Int32 dwType</param>
		/// <param name="pImeParentMenu">__MIDL___MIDL_itf_mshtml_0001_0042_0010 pImeParentMenu</param>
		/// <param name="pImeMenu">__MIDL___MIDL_itf_mshtml_0001_0042_0010 pImeMenu</param>
		/// <param name="dwSize">Int32 dwSize</param>
		/// <param name="pdwResult">Int32 pdwResult</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetImeMenuItemsA(Int32 hIMC, Int32 dwFlags, Int32 dwType, __MIDL___MIDL_itf_mshtml_0001_0042_0010 pImeParentMenu, out __MIDL___MIDL_itf_mshtml_0001_0042_0010 pImeMenu, Int32 dwSize, out Int32 pdwResult);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="dwType">Int32 dwType</param>
		/// <param name="pImeParentMenu">__MIDL___MIDL_itf_mshtml_0001_0042_0011 pImeParentMenu</param>
		/// <param name="pImeMenu">__MIDL___MIDL_itf_mshtml_0001_0042_0011 pImeMenu</param>
		/// <param name="dwSize">Int32 dwSize</param>
		/// <param name="pdwResult">Int32 pdwResult</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 GetImeMenuItemsW(Int32 hIMC, Int32 dwFlags, Int32 dwType, __MIDL___MIDL_itf_mshtml_0001_0042_0011 pImeParentMenu, out __MIDL___MIDL_itf_mshtml_0001_0042_0011 pImeMenu, Int32 dwSize, out Int32 pdwResult);

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="idThread">Int32 idThread</param>
		/// <param name="ppEnum">NetOffice.MSHTMLApi.IEnumInputContext ppEnum</param>
		[SupportByVersion("MSHTML", 4)]
		Int32 EnumInputContext(Int32 idThread, out NetOffice.MSHTMLApi.IEnumInputContext ppEnum);

		#endregion
	}
}
