using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.MSHTMLApi
{
	///<summary>
	/// Interface IActiveIMMApp 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IActiveIMMApp : COMObject
	{
		#pragma warning disable
		#region Type Information

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(IActiveIMMApp);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IActiveIMMApp(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IActiveIMMApp(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IActiveIMMApp(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IActiveIMMApp(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IActiveIMMApp(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IActiveIMMApp() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IActiveIMMApp(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="hIME">Int32 hIME</param>
		/// <param name="phPrev">Int32 phPrev</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 AssociateContext(_RemotableHandle hWnd, Int32 hIME, out Int32 phPrev)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			phPrev = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hWnd, hIME, phPrev);
			object returnItem = Invoker.MethodReturn(this, "AssociateContext", paramsArray);
			phPrev = (Int32)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="dwMode">Int32 dwMode</param>
		/// <param name="pData">__MIDL___MIDL_itf_mshtml_0001_0042_0001 pData</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 ConfigureIMEA(object hKL, _RemotableHandle hWnd, Int32 dwMode, __MIDL___MIDL_itf_mshtml_0001_0042_0001 pData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, hWnd, dwMode, pData);
			object returnItem = Invoker.MethodReturn(this, "ConfigureIMEA", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="dwMode">Int32 dwMode</param>
		/// <param name="pData">__MIDL___MIDL_itf_mshtml_0001_0042_0002 pData</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 ConfigureIMEW(object hKL, _RemotableHandle hWnd, Int32 dwMode, __MIDL___MIDL_itf_mshtml_0001_0042_0002 pData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, hWnd, dwMode, pData);
			object returnItem = Invoker.MethodReturn(this, "ConfigureIMEW", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="phIMC">Int32 phIMC</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 CreateContext(out Int32 phIMC)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			phIMC = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(phIMC);
			object returnItem = Invoker.MethodReturn(this, "CreateContext", paramsArray);
			phIMC = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIME">Int32 hIME</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 DestroyContext(Int32 hIME)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hIME);
			object returnItem = Invoker.MethodReturn(this, "DestroyContext", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="szReading">string szReading</param>
		/// <param name="dwStyle">Int32 dwStyle</param>
		/// <param name="szRegister">string szRegister</param>
		/// <param name="pData">object pData</param>
		/// <param name="pEnum">NetOffice.MSHTMLApi.IEnumRegisterWordA pEnum</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 EnumRegisterWordA(object hKL, string szReading, Int32 dwStyle, string szRegister, object pData, out NetOffice.MSHTMLApi.IEnumRegisterWordA pEnum)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,true);
			pEnum = null;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, szReading, dwStyle, szRegister, pData, pEnum);
			object returnItem = Invoker.MethodReturn(this, "EnumRegisterWordA", paramsArray);
			pEnum = (NetOffice.MSHTMLApi.IEnumRegisterWordA)paramsArray[5];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="szReading">string szReading</param>
		/// <param name="dwStyle">Int32 dwStyle</param>
		/// <param name="szRegister">string szRegister</param>
		/// <param name="pData">object pData</param>
		/// <param name="pEnum">NetOffice.MSHTMLApi.IEnumRegisterWordW pEnum</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 EnumRegisterWordW(object hKL, string szReading, Int32 dwStyle, string szRegister, object pData, out NetOffice.MSHTMLApi.IEnumRegisterWordW pEnum)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,true);
			pEnum = null;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, szReading, dwStyle, szRegister, pData, pEnum);
			object returnItem = Invoker.MethodReturn(this, "EnumRegisterWordW", paramsArray);
			pEnum = (NetOffice.MSHTMLApi.IEnumRegisterWordW)paramsArray[5];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="uEscape">UIntPtr uEscape</param>
		/// <param name="pData">object pData</param>
		/// <param name="plResult">Int32 plResult</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 EscapeA(object hKL, Int32 hIMC, UIntPtr uEscape, object pData, out Int32 plResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true);
			plResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, hIMC, uEscape, pData, plResult);
			object returnItem = Invoker.MethodReturn(this, "EscapeA", paramsArray);
			plResult = (Int32)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="uEscape">UIntPtr uEscape</param>
		/// <param name="pData">object pData</param>
		/// <param name="plResult">Int32 plResult</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 EscapeW(object hKL, Int32 hIMC, UIntPtr uEscape, object pData, out Int32 plResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true);
			plResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, hIMC, uEscape, pData, plResult);
			object returnItem = Invoker.MethodReturn(this, "EscapeW", paramsArray);
			plResult = (Int32)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="pCandList">__MIDL___MIDL_itf_mshtml_0001_0042_0007 pCandList</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetCandidateListA(Int32 hIMC, Int32 dwIndex, UIntPtr uBufLen, out __MIDL___MIDL_itf_mshtml_0001_0042_0007 pCandList, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true);
			pCandList = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0007();
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwIndex, uBufLen, pCandList, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetCandidateListA", paramsArray);
			pCandList = (__MIDL___MIDL_itf_mshtml_0001_0042_0007)paramsArray[3];
			puCopied = (UIntPtr)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="pCandList">__MIDL___MIDL_itf_mshtml_0001_0042_0007 pCandList</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetCandidateListW(Int32 hIMC, Int32 dwIndex, UIntPtr uBufLen, out __MIDL___MIDL_itf_mshtml_0001_0042_0007 pCandList, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true);
			pCandList = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0007();
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwIndex, uBufLen, pCandList, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetCandidateListW", paramsArray);
			pCandList = (__MIDL___MIDL_itf_mshtml_0001_0042_0007)paramsArray[3];
			puCopied = (UIntPtr)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pdwListSize">Int32 pdwListSize</param>
		/// <param name="pdwBufLen">Int32 pdwBufLen</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetCandidateListCountA(Int32 hIMC, out Int32 pdwListSize, out Int32 pdwBufLen)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true);
			pdwListSize = 0;
			pdwBufLen = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, pdwListSize, pdwBufLen);
			object returnItem = Invoker.MethodReturn(this, "GetCandidateListCountA", paramsArray);
			pdwListSize = (Int32)paramsArray[1];
			pdwBufLen = (Int32)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pdwListSize">Int32 pdwListSize</param>
		/// <param name="pdwBufLen">Int32 pdwBufLen</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetCandidateListCountW(Int32 hIMC, out Int32 pdwListSize, out Int32 pdwBufLen)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true);
			pdwListSize = 0;
			pdwBufLen = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, pdwListSize, pdwBufLen);
			object returnItem = Invoker.MethodReturn(this, "GetCandidateListCountW", paramsArray);
			pdwListSize = (Int32)paramsArray[1];
			pdwBufLen = (Int32)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="pCandidate">__MIDL___MIDL_itf_mshtml_0001_0042_0005 pCandidate</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetCandidateWindow(Int32 hIMC, Int32 dwIndex, out __MIDL___MIDL_itf_mshtml_0001_0042_0005 pCandidate)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			pCandidate = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0005();
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwIndex, pCandidate);
			object returnItem = Invoker.MethodReturn(this, "GetCandidateWindow", paramsArray);
			pCandidate = (__MIDL___MIDL_itf_mshtml_0001_0042_0005)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="plf">__MIDL___MIDL_itf_mshtml_0001_0042_0003 plf</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetCompositionFontA(Int32 hIMC, out __MIDL___MIDL_itf_mshtml_0001_0042_0003 plf)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			plf = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0003();
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, plf);
			object returnItem = Invoker.MethodReturn(this, "GetCompositionFontA", paramsArray);
			plf = (__MIDL___MIDL_itf_mshtml_0001_0042_0003)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="plf">__MIDL___MIDL_itf_mshtml_0001_0042_0004 plf</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetCompositionFontW(Int32 hIMC, out __MIDL___MIDL_itf_mshtml_0001_0042_0004 plf)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			plf = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0004();
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, plf);
			object returnItem = Invoker.MethodReturn(this, "GetCompositionFontW", paramsArray);
			plf = (__MIDL___MIDL_itf_mshtml_0001_0042_0004)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="dwBufLen">Int32 dwBufLen</param>
		/// <param name="plCopied">Int32 plCopied</param>
		/// <param name="pBuf">object pBuf</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetCompositionStringA(Int32 hIMC, Int32 dwIndex, Int32 dwBufLen, out Int32 plCopied, out object pBuf)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true);
			plCopied = 0;
			pBuf = null;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwIndex, dwBufLen, plCopied, pBuf);
			object returnItem = Invoker.MethodReturn(this, "GetCompositionStringA", paramsArray);
			plCopied = (Int32)paramsArray[3];
			pBuf = (object)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="dwBufLen">Int32 dwBufLen</param>
		/// <param name="plCopied">Int32 plCopied</param>
		/// <param name="pBuf">object pBuf</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetCompositionStringW(Int32 hIMC, Int32 dwIndex, Int32 dwBufLen, out Int32 plCopied, out object pBuf)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true);
			plCopied = 0;
			pBuf = null;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwIndex, dwBufLen, plCopied, pBuf);
			object returnItem = Invoker.MethodReturn(this, "GetCompositionStringW", paramsArray);
			plCopied = (Int32)paramsArray[3];
			pBuf = (object)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pCompForm">__MIDL___MIDL_itf_mshtml_0001_0042_0006 pCompForm</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetCompositionWindow(Int32 hIMC, out __MIDL___MIDL_itf_mshtml_0001_0042_0006 pCompForm)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pCompForm = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0006();
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, pCompForm);
			object returnItem = Invoker.MethodReturn(this, "GetCompositionWindow", paramsArray);
			pCompForm = (__MIDL___MIDL_itf_mshtml_0001_0042_0006)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="phIMC">Int32 phIMC</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetContext(_RemotableHandle hWnd, out Int32 phIMC)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			phIMC = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hWnd, phIMC);
			object returnItem = Invoker.MethodReturn(this, "GetContext", paramsArray);
			phIMC = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pSrc">string pSrc</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="uFlag">UIntPtr uFlag</param>
		/// <param name="pDst">__MIDL___MIDL_itf_mshtml_0001_0042_0007 pDst</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetConversionListA(object hKL, Int32 hIMC, string pSrc, UIntPtr uBufLen, UIntPtr uFlag, out __MIDL___MIDL_itf_mshtml_0001_0042_0007 pDst, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,true,true);
			pDst = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0007();
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, hIMC, pSrc, uBufLen, uFlag, pDst, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetConversionListA", paramsArray);
			pDst = (__MIDL___MIDL_itf_mshtml_0001_0042_0007)paramsArray[5];
			puCopied = (UIntPtr)paramsArray[6];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pSrc">string pSrc</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="uFlag">UIntPtr uFlag</param>
		/// <param name="pDst">__MIDL___MIDL_itf_mshtml_0001_0042_0007 pDst</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetConversionListW(object hKL, Int32 hIMC, string pSrc, UIntPtr uBufLen, UIntPtr uFlag, out __MIDL___MIDL_itf_mshtml_0001_0042_0007 pDst, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,true,true);
			pDst = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0007();
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, hIMC, pSrc, uBufLen, uFlag, pDst, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetConversionListW", paramsArray);
			pDst = (__MIDL___MIDL_itf_mshtml_0001_0042_0007)paramsArray[5];
			puCopied = (UIntPtr)paramsArray[6];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pfdwConversion">Int32 pfdwConversion</param>
		/// <param name="pfdwSentence">Int32 pfdwSentence</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetConversionStatus(Int32 hIMC, out Int32 pfdwConversion, out Int32 pfdwSentence)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true);
			pfdwConversion = 0;
			pfdwSentence = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, pfdwConversion, pfdwSentence);
			object returnItem = Invoker.MethodReturn(this, "GetConversionStatus", paramsArray);
			pfdwConversion = (Int32)paramsArray[1];
			pfdwSentence = (Int32)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="phDefWnd">_RemotableHandle phDefWnd</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetDefaultIMEWnd(_RemotableHandle hWnd, out _RemotableHandle phDefWnd)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			phDefWnd = new NetOffice.MSHTMLApi._RemotableHandle();
			object[] paramsArray = Invoker.ValidateParamsArray(hWnd, phDefWnd);
			object returnItem = Invoker.MethodReturn(this, "GetDefaultIMEWnd", paramsArray);
			phDefWnd = (_RemotableHandle)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="szDescription">string szDescription</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetDescriptionA(object hKL, UIntPtr uBufLen, out string szDescription, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true);
			szDescription = string.Empty;
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, uBufLen, szDescription, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetDescriptionA", paramsArray);
			szDescription = (string)paramsArray[2];
			puCopied = (UIntPtr)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="szDescription">string szDescription</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetDescriptionW(object hKL, UIntPtr uBufLen, out string szDescription, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true);
			szDescription = string.Empty;
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, uBufLen, szDescription, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetDescriptionW", paramsArray);
			szDescription = (string)paramsArray[2];
			puCopied = (UIntPtr)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="dwBufLen">Int32 dwBufLen</param>
		/// <param name="pBuf">string pBuf</param>
		/// <param name="pdwResult">Int32 pdwResult</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetGuideLineA(Int32 hIMC, Int32 dwIndex, Int32 dwBufLen, out string pBuf, out Int32 pdwResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true);
			pBuf = string.Empty;
			pdwResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwIndex, dwBufLen, pBuf, pdwResult);
			object returnItem = Invoker.MethodReturn(this, "GetGuideLineA", paramsArray);
			pBuf = (string)paramsArray[3];
			pdwResult = (Int32)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="dwBufLen">Int32 dwBufLen</param>
		/// <param name="pBuf">string pBuf</param>
		/// <param name="pdwResult">Int32 pdwResult</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetGuideLineW(Int32 hIMC, Int32 dwIndex, Int32 dwBufLen, out string pBuf, out Int32 pdwResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true);
			pBuf = string.Empty;
			pdwResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwIndex, dwBufLen, pBuf, pdwResult);
			object returnItem = Invoker.MethodReturn(this, "GetGuideLineW", paramsArray);
			pBuf = (string)paramsArray[3];
			pdwResult = (Int32)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="szFileName">string szFileName</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetIMEFileNameA(object hKL, UIntPtr uBufLen, out string szFileName, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true);
			szFileName = string.Empty;
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, uBufLen, szFileName, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetIMEFileNameA", paramsArray);
			szFileName = (string)paramsArray[2];
			puCopied = (UIntPtr)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="szFileName">string szFileName</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetIMEFileNameW(object hKL, UIntPtr uBufLen, out string szFileName, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true);
			szFileName = string.Empty;
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, uBufLen, szFileName, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetIMEFileNameW", paramsArray);
			szFileName = (string)paramsArray[2];
			puCopied = (UIntPtr)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetOpenStatus(Int32 hIMC)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC);
			object returnItem = Invoker.MethodReturn(this, "GetOpenStatus", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="fdwIndex">Int32 fdwIndex</param>
		/// <param name="pdwProperty">Int32 pdwProperty</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetProperty(object hKL, Int32 fdwIndex, out Int32 pdwProperty)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			pdwProperty = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, fdwIndex, pdwProperty);
			object returnItem = Invoker.MethodReturn(this, "GetProperty", paramsArray);
			pdwProperty = (Int32)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="nItem">UIntPtr nItem</param>
		/// <param name="pStyleBuf">__MIDL___MIDL_itf_mshtml_0001_0042_0008 pStyleBuf</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetRegisterWordStyleA(object hKL, UIntPtr nItem, out __MIDL___MIDL_itf_mshtml_0001_0042_0008 pStyleBuf, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true);
			pStyleBuf = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0008();
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, nItem, pStyleBuf, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetRegisterWordStyleA", paramsArray);
			pStyleBuf = (__MIDL___MIDL_itf_mshtml_0001_0042_0008)paramsArray[2];
			puCopied = (UIntPtr)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="nItem">UIntPtr nItem</param>
		/// <param name="pStyleBuf">__MIDL___MIDL_itf_mshtml_0001_0042_0009 pStyleBuf</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetRegisterWordStyleW(object hKL, UIntPtr nItem, out __MIDL___MIDL_itf_mshtml_0001_0042_0009 pStyleBuf, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true);
			pStyleBuf = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0009();
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, nItem, pStyleBuf, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetRegisterWordStyleW", paramsArray);
			pStyleBuf = (__MIDL___MIDL_itf_mshtml_0001_0042_0009)paramsArray[2];
			puCopied = (UIntPtr)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pptPos">tagPOINT pptPos</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetStatusWindowPos(Int32 hIMC, out tagPOINT pptPos)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pptPos = new NetOffice.MSHTMLApi.tagPOINT();
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, pptPos);
			object returnItem = Invoker.MethodReturn(this, "GetStatusWindowPos", paramsArray);
			pptPos = (tagPOINT)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="puVirtualKey">UIntPtr puVirtualKey</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetVirtualKey(_RemotableHandle hWnd, out UIntPtr puVirtualKey)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			puVirtualKey = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hWnd, puVirtualKey);
			object returnItem = Invoker.MethodReturn(this, "GetVirtualKey", paramsArray);
			puVirtualKey = (UIntPtr)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="szIMEFileName">string szIMEFileName</param>
		/// <param name="szLayoutText">string szLayoutText</param>
		/// <param name="phKL">object phKL</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 InstallIMEA(string szIMEFileName, string szLayoutText, out object phKL)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			phKL = null;
			object[] paramsArray = Invoker.ValidateParamsArray(szIMEFileName, szLayoutText, phKL);
			object returnItem = Invoker.MethodReturn(this, "InstallIMEA", paramsArray);
			phKL = (object)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="szIMEFileName">string szIMEFileName</param>
		/// <param name="szLayoutText">string szLayoutText</param>
		/// <param name="phKL">object phKL</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 InstallIMEW(string szIMEFileName, string szLayoutText, out object phKL)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			phKL = null;
			object[] paramsArray = Invoker.ValidateParamsArray(szIMEFileName, szLayoutText, phKL);
			object returnItem = Invoker.MethodReturn(this, "InstallIMEW", paramsArray);
			phKL = (object)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 IsIME(object hKL)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hKL);
			object returnItem = Invoker.MethodReturn(this, "IsIME", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hWndIME">_RemotableHandle hWndIME</param>
		/// <param name="msg">UIntPtr msg</param>
		/// <param name="wParam">Int32 wParam</param>
		/// <param name="lParam">Int32 lParam</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 IsUIMessageA(_RemotableHandle hWndIME, UIntPtr msg, Int32 wParam, Int32 lParam)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hWndIME, msg, wParam, lParam);
			object returnItem = Invoker.MethodReturn(this, "IsUIMessageA", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hWndIME">_RemotableHandle hWndIME</param>
		/// <param name="msg">UIntPtr msg</param>
		/// <param name="wParam">Int32 wParam</param>
		/// <param name="lParam">Int32 lParam</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 IsUIMessageW(_RemotableHandle hWndIME, UIntPtr msg, Int32 wParam, Int32 lParam)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hWndIME, msg, wParam, lParam);
			object returnItem = Invoker.MethodReturn(this, "IsUIMessageW", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwAction">Int32 dwAction</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="dwValue">Int32 dwValue</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 NotifyIME(Int32 hIMC, Int32 dwAction, Int32 dwIndex, Int32 dwValue)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwAction, dwIndex, dwValue);
			object returnItem = Invoker.MethodReturn(this, "NotifyIME", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="szReading">string szReading</param>
		/// <param name="dwStyle">Int32 dwStyle</param>
		/// <param name="szRegister">string szRegister</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 RegisterWordA(object hKL, string szReading, Int32 dwStyle, string szRegister)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, szReading, dwStyle, szRegister);
			object returnItem = Invoker.MethodReturn(this, "RegisterWordA", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="szReading">string szReading</param>
		/// <param name="dwStyle">Int32 dwStyle</param>
		/// <param name="szRegister">string szRegister</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 RegisterWordW(object hKL, string szReading, Int32 dwStyle, string szRegister)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, szReading, dwStyle, szRegister);
			object returnItem = Invoker.MethodReturn(this, "RegisterWordW", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="hIMC">Int32 hIMC</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 ReleaseContext(_RemotableHandle hWnd, Int32 hIMC)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hWnd, hIMC);
			object returnItem = Invoker.MethodReturn(this, "ReleaseContext", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pCandidate">__MIDL___MIDL_itf_mshtml_0001_0042_0005 pCandidate</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 SetCandidateWindow(Int32 hIMC, __MIDL___MIDL_itf_mshtml_0001_0042_0005 pCandidate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, pCandidate);
			object returnItem = Invoker.MethodReturn(this, "SetCandidateWindow", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="plf">__MIDL___MIDL_itf_mshtml_0001_0042_0003 plf</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 SetCompositionFontA(Int32 hIMC, __MIDL___MIDL_itf_mshtml_0001_0042_0003 plf)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, plf);
			object returnItem = Invoker.MethodReturn(this, "SetCompositionFontA", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="plf">__MIDL___MIDL_itf_mshtml_0001_0042_0004 plf</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 SetCompositionFontW(Int32 hIMC, __MIDL___MIDL_itf_mshtml_0001_0042_0004 plf)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, plf);
			object returnItem = Invoker.MethodReturn(this, "SetCompositionFontW", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="pComp">object pComp</param>
		/// <param name="dwCompLen">Int32 dwCompLen</param>
		/// <param name="pRead">object pRead</param>
		/// <param name="dwReadLen">Int32 dwReadLen</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 SetCompositionStringA(Int32 hIMC, Int32 dwIndex, object pComp, Int32 dwCompLen, object pRead, Int32 dwReadLen)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwIndex, pComp, dwCompLen, pRead, dwReadLen);
			object returnItem = Invoker.MethodReturn(this, "SetCompositionStringA", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="pComp">object pComp</param>
		/// <param name="dwCompLen">Int32 dwCompLen</param>
		/// <param name="pRead">object pRead</param>
		/// <param name="dwReadLen">Int32 dwReadLen</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 SetCompositionStringW(Int32 hIMC, Int32 dwIndex, object pComp, Int32 dwCompLen, object pRead, Int32 dwReadLen)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwIndex, pComp, dwCompLen, pRead, dwReadLen);
			object returnItem = Invoker.MethodReturn(this, "SetCompositionStringW", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pCompForm">__MIDL___MIDL_itf_mshtml_0001_0042_0006 pCompForm</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 SetCompositionWindow(Int32 hIMC, __MIDL___MIDL_itf_mshtml_0001_0042_0006 pCompForm)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, pCompForm);
			object returnItem = Invoker.MethodReturn(this, "SetCompositionWindow", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="fdwConversion">Int32 fdwConversion</param>
		/// <param name="fdwSentence">Int32 fdwSentence</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 SetConversionStatus(Int32 hIMC, Int32 fdwConversion, Int32 fdwSentence)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, fdwConversion, fdwSentence);
			object returnItem = Invoker.MethodReturn(this, "SetConversionStatus", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="fOpen">Int32 fOpen</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 SetOpenStatus(Int32 hIMC, Int32 fOpen)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, fOpen);
			object returnItem = Invoker.MethodReturn(this, "SetOpenStatus", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pptPos">tagPOINT pptPos</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 SetStatusWindowPos(Int32 hIMC, tagPOINT pptPos)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, pptPos);
			object returnItem = Invoker.MethodReturn(this, "SetStatusWindowPos", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="dwHotKeyID">Int32 dwHotKeyID</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 SimulateHotKey(_RemotableHandle hWnd, Int32 dwHotKeyID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hWnd, dwHotKeyID);
			object returnItem = Invoker.MethodReturn(this, "SimulateHotKey", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="szReading">string szReading</param>
		/// <param name="dwStyle">Int32 dwStyle</param>
		/// <param name="szUnregister">string szUnregister</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 UnregisterWordA(object hKL, string szReading, Int32 dwStyle, string szUnregister)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, szReading, dwStyle, szUnregister);
			object returnItem = Invoker.MethodReturn(this, "UnregisterWordA", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="szReading">string szReading</param>
		/// <param name="dwStyle">Int32 dwStyle</param>
		/// <param name="szUnregister">string szUnregister</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 UnregisterWordW(object hKL, string szReading, Int32 dwStyle, string szUnregister)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, szReading, dwStyle, szUnregister);
			object returnItem = Invoker.MethodReturn(this, "UnregisterWordW", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="fRestoreLayout">Int32 fRestoreLayout</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 Activate(Int32 fRestoreLayout)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fRestoreLayout);
			object returnItem = Invoker.MethodReturn(this, "Activate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 Deactivate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Deactivate", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="msg">UIntPtr msg</param>
		/// <param name="wParam">Int32 wParam</param>
		/// <param name="lParam">Int32 lParam</param>
		/// <param name="plResult">Int32 plResult</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 OnDefWindowProc(_RemotableHandle hWnd, UIntPtr msg, Int32 wParam, Int32 lParam, out Int32 plResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true);
			plResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hWnd, msg, wParam, lParam, plResult);
			object returnItem = Invoker.MethodReturn(this, "OnDefWindowProc", paramsArray);
			plResult = (Int32)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="aaClassList">Int16 aaClassList</param>
		/// <param name="uSize">UIntPtr uSize</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 FilterClientWindows(Int16 aaClassList, UIntPtr uSize)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(aaClassList, uSize);
			object returnItem = Invoker.MethodReturn(this, "FilterClientWindows", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="uCodePage">UIntPtr uCodePage</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetCodePageA(object hKL, out UIntPtr uCodePage)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			uCodePage = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, uCodePage);
			object returnItem = Invoker.MethodReturn(this, "GetCodePageA", paramsArray);
			uCodePage = (UIntPtr)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="plid">Int16 plid</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetLangId(object hKL, out Int16 plid)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			plid = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, plid);
			object returnItem = Invoker.MethodReturn(this, "GetLangId", paramsArray);
			plid = (Int16)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 AssociateContextEx(_RemotableHandle hWnd, Int32 hIMC, Int32 dwFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hWnd, hIMC, dwFlags);
			object returnItem = Invoker.MethodReturn(this, "AssociateContextEx", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="idThread">Int32 idThread</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 DisableIME(Int32 idThread)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(idThread);
			object returnItem = Invoker.MethodReturn(this, "DisableIME", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="dwType">Int32 dwType</param>
		/// <param name="pImeParentMenu">__MIDL___MIDL_itf_mshtml_0001_0042_0010 pImeParentMenu</param>
		/// <param name="pImeMenu">__MIDL___MIDL_itf_mshtml_0001_0042_0010 pImeMenu</param>
		/// <param name="dwSize">Int32 dwSize</param>
		/// <param name="pdwResult">Int32 pdwResult</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetImeMenuItemsA(Int32 hIMC, Int32 dwFlags, Int32 dwType, __MIDL___MIDL_itf_mshtml_0001_0042_0010 pImeParentMenu, out __MIDL___MIDL_itf_mshtml_0001_0042_0010 pImeMenu, Int32 dwSize, out Int32 pdwResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true,false,true);
			pImeMenu = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0010();
			pdwResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwFlags, dwType, pImeParentMenu, pImeMenu, dwSize, pdwResult);
			object returnItem = Invoker.MethodReturn(this, "GetImeMenuItemsA", paramsArray);
			pImeMenu = (__MIDL___MIDL_itf_mshtml_0001_0042_0010)paramsArray[4];
			pdwResult = (Int32)paramsArray[6];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="dwType">Int32 dwType</param>
		/// <param name="pImeParentMenu">__MIDL___MIDL_itf_mshtml_0001_0042_0011 pImeParentMenu</param>
		/// <param name="pImeMenu">__MIDL___MIDL_itf_mshtml_0001_0042_0011 pImeMenu</param>
		/// <param name="dwSize">Int32 dwSize</param>
		/// <param name="pdwResult">Int32 pdwResult</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetImeMenuItemsW(Int32 hIMC, Int32 dwFlags, Int32 dwType, __MIDL___MIDL_itf_mshtml_0001_0042_0011 pImeParentMenu, out __MIDL___MIDL_itf_mshtml_0001_0042_0011 pImeMenu, Int32 dwSize, out Int32 pdwResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true,false,true);
			pImeMenu = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0011();
			pdwResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwFlags, dwType, pImeParentMenu, pImeMenu, dwSize, pdwResult);
			object returnItem = Invoker.MethodReturn(this, "GetImeMenuItemsW", paramsArray);
			pImeMenu = (__MIDL___MIDL_itf_mshtml_0001_0042_0011)paramsArray[4];
			pdwResult = (Int32)paramsArray[6];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="idThread">Int32 idThread</param>
		/// <param name="ppEnum">NetOffice.MSHTMLApi.IEnumInputContext ppEnum</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 EnumInputContext(Int32 idThread, out NetOffice.MSHTMLApi.IEnumInputContext ppEnum)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			ppEnum = null;
			object[] paramsArray = Invoker.ValidateParamsArray(idThread, ppEnum);
			object returnItem = Invoker.MethodReturn(this, "EnumInputContext", paramsArray);
			ppEnum = (NetOffice.MSHTMLApi.IEnumInputContext)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}