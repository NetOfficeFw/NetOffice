using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IActiveIMMApp 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IActiveIMMApp : COMObject, NetOffice.MSHTMLApi.IActiveIMMApp
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.MSHTMLApi.IActiveIMMApp);
                return _contractType;
            }
        }
        private static Type _contractType;


		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

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
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IActiveIMMApp() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="hIME">Int32 hIME</param>
		/// <param name="phPrev">Int32 phPrev</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 AssociateContext(_RemotableHandle hWnd, Int32 hIME, out Int32 phPrev)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			phPrev = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hWnd, hIME, phPrev);
			object returnItem = Invoker.MethodReturn(this, "AssociateContext", paramsArray, modifiers);
			phPrev = (Int32)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="dwMode">Int32 dwMode</param>
		/// <param name="pData">__MIDL___MIDL_itf_mshtml_0001_0042_0001 pData</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 ConfigureIMEA(object hKL, _RemotableHandle hWnd, Int32 dwMode, __MIDL___MIDL_itf_mshtml_0001_0042_0001 pData)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ConfigureIMEA", hKL, hWnd, dwMode, pData);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="dwMode">Int32 dwMode</param>
		/// <param name="pData">__MIDL___MIDL_itf_mshtml_0001_0042_0002 pData</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 ConfigureIMEW(object hKL, _RemotableHandle hWnd, Int32 dwMode, __MIDL___MIDL_itf_mshtml_0001_0042_0002 pData)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ConfigureIMEW", hKL, hWnd, dwMode, pData);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="phIMC">Int32 phIMC</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 CreateContext(out Int32 phIMC)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			phIMC = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(phIMC);
			object returnItem = Invoker.MethodReturn(this, "CreateContext", paramsArray, modifiers);
			phIMC = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIME">Int32 hIME</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 DestroyContext(Int32 hIME)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DestroyContext", hIME);
		}

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
		public virtual Int32 EnumRegisterWordA(object hKL, string szReading, Int32 dwStyle, string szRegister, object pData, out NetOffice.MSHTMLApi.IEnumRegisterWordA pEnum)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,true);
			pEnum = null;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, szReading, dwStyle, szRegister, pData, pEnum);
			object returnItem = Invoker.MethodReturn(this, "EnumRegisterWordA", paramsArray, modifiers);
            if (paramsArray[5] is MarshalByRefObject)
                pEnum = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IEnumRegisterWordA>(this, paramsArray[5], typeof(NetOffice.MSHTMLApi.IEnumRegisterWordA));
            else
                pEnum = null;
            return NetRuntimeSystem.Convert.ToInt32(returnItem);            
        }

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
		public virtual Int32 EnumRegisterWordW(object hKL, string szReading, Int32 dwStyle, string szRegister, object pData, out NetOffice.MSHTMLApi.IEnumRegisterWordW pEnum)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,true);
			pEnum = null;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, szReading, dwStyle, szRegister, pData, pEnum);
			object returnItem = Invoker.MethodReturn(this, "EnumRegisterWordW", paramsArray, modifiers);
            if (paramsArray[5] is MarshalByRefObject)
                pEnum = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IEnumRegisterWordW>(this, paramsArray[5], typeof(NetOffice.MSHTMLApi.IEnumRegisterWordW));
            else
                pEnum = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);            
        }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="uEscape">UIntPtr uEscape</param>
		/// <param name="pData">object pData</param>
		/// <param name="plResult">Int32 plResult</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 EscapeA(object hKL, Int32 hIMC, UIntPtr uEscape, object pData, out Int32 plResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true);
			plResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, hIMC, uEscape, pData, plResult);
			object returnItem = Invoker.MethodReturn(this, "EscapeA", paramsArray, modifiers);
            plResult = (Int32)paramsArray[4];
            return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="uEscape">UIntPtr uEscape</param>
		/// <param name="pData">object pData</param>
		/// <param name="plResult">Int32 plResult</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 EscapeW(object hKL, Int32 hIMC, UIntPtr uEscape, object pData, out Int32 plResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true);
			plResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, hIMC, uEscape, pData, plResult);
			object returnItem = Invoker.MethodReturn(this, "EscapeW", paramsArray, modifiers);
			plResult = (Int32)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="pCandList">__MIDL___MIDL_itf_mshtml_0001_0042_0007 pCandList</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetCandidateListA(Int32 hIMC, Int32 dwIndex, UIntPtr uBufLen, out __MIDL___MIDL_itf_mshtml_0001_0042_0007 pCandList, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true);
			pCandList = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0007();
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwIndex, uBufLen, pCandList, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetCandidateListA", paramsArray, modifiers);
			pCandList = (__MIDL___MIDL_itf_mshtml_0001_0042_0007)paramsArray[3];
			puCopied = (UIntPtr)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="pCandList">__MIDL___MIDL_itf_mshtml_0001_0042_0007 pCandList</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetCandidateListW(Int32 hIMC, Int32 dwIndex, UIntPtr uBufLen, out __MIDL___MIDL_itf_mshtml_0001_0042_0007 pCandList, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true);
			pCandList = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0007();
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwIndex, uBufLen, pCandList, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetCandidateListW", paramsArray, modifiers);
			pCandList = (__MIDL___MIDL_itf_mshtml_0001_0042_0007)paramsArray[3];
			puCopied = (UIntPtr)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pdwListSize">Int32 pdwListSize</param>
		/// <param name="pdwBufLen">Int32 pdwBufLen</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetCandidateListCountA(Int32 hIMC, out Int32 pdwListSize, out Int32 pdwBufLen)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true);
			pdwListSize = 0;
			pdwBufLen = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, pdwListSize, pdwBufLen);
			object returnItem = Invoker.MethodReturn(this, "GetCandidateListCountA", paramsArray, modifiers);
			pdwListSize = (Int32)paramsArray[1];
			pdwBufLen = (Int32)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pdwListSize">Int32 pdwListSize</param>
		/// <param name="pdwBufLen">Int32 pdwBufLen</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetCandidateListCountW(Int32 hIMC, out Int32 pdwListSize, out Int32 pdwBufLen)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true);
			pdwListSize = 0;
			pdwBufLen = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, pdwListSize, pdwBufLen);
			object returnItem = Invoker.MethodReturn(this, "GetCandidateListCountW", paramsArray, modifiers);
			pdwListSize = (Int32)paramsArray[1];
			pdwBufLen = (Int32)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="pCandidate">__MIDL___MIDL_itf_mshtml_0001_0042_0005 pCandidate</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetCandidateWindow(Int32 hIMC, Int32 dwIndex, out __MIDL___MIDL_itf_mshtml_0001_0042_0005 pCandidate)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			pCandidate = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0005();
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwIndex, pCandidate);
			object returnItem = Invoker.MethodReturn(this, "GetCandidateWindow", paramsArray, modifiers);
			pCandidate = (__MIDL___MIDL_itf_mshtml_0001_0042_0005)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="plf">__MIDL___MIDL_itf_mshtml_0001_0042_0003 plf</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetCompositionFontA(Int32 hIMC, out __MIDL___MIDL_itf_mshtml_0001_0042_0003 plf)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			plf = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0003();
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, plf);
			object returnItem = Invoker.MethodReturn(this, "GetCompositionFontA", paramsArray, modifiers);
			plf = (__MIDL___MIDL_itf_mshtml_0001_0042_0003)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="plf">__MIDL___MIDL_itf_mshtml_0001_0042_0004 plf</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetCompositionFontW(Int32 hIMC, out __MIDL___MIDL_itf_mshtml_0001_0042_0004 plf)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			plf = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0004();
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, plf);
			object returnItem = Invoker.MethodReturn(this, "GetCompositionFontW", paramsArray, modifiers);
			plf = (__MIDL___MIDL_itf_mshtml_0001_0042_0004)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="dwBufLen">Int32 dwBufLen</param>
		/// <param name="plCopied">Int32 plCopied</param>
		/// <param name="pBuf">object pBuf</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetCompositionStringA(Int32 hIMC, Int32 dwIndex, Int32 dwBufLen, out Int32 plCopied, out object pBuf)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true);
			plCopied = 0;
			pBuf = null;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwIndex, dwBufLen, plCopied, pBuf);
			object returnItem = Invoker.MethodReturn(this, "GetCompositionStringA", paramsArray, modifiers);
			plCopied = (Int32)paramsArray[3];
			pBuf = (object)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="dwBufLen">Int32 dwBufLen</param>
		/// <param name="plCopied">Int32 plCopied</param>
		/// <param name="pBuf">object pBuf</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetCompositionStringW(Int32 hIMC, Int32 dwIndex, Int32 dwBufLen, out Int32 plCopied, out object pBuf)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true);
			plCopied = 0;
			pBuf = null;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwIndex, dwBufLen, plCopied, pBuf);
			object returnItem = Invoker.MethodReturn(this, "GetCompositionStringW", paramsArray, modifiers);
			plCopied = (Int32)paramsArray[3];
			pBuf = (object)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pCompForm">__MIDL___MIDL_itf_mshtml_0001_0042_0006 pCompForm</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetCompositionWindow(Int32 hIMC, out __MIDL___MIDL_itf_mshtml_0001_0042_0006 pCompForm)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pCompForm = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0006();
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, pCompForm);
			object returnItem = Invoker.MethodReturn(this, "GetCompositionWindow", paramsArray, modifiers);
			pCompForm = (__MIDL___MIDL_itf_mshtml_0001_0042_0006)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="phIMC">Int32 phIMC</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetContext(_RemotableHandle hWnd, out Int32 phIMC)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			phIMC = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hWnd, phIMC);
			object returnItem = Invoker.MethodReturn(this, "GetContext", paramsArray, modifiers);
			phIMC = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

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
		public virtual Int32 GetConversionListA(object hKL, Int32 hIMC, string pSrc, UIntPtr uBufLen, UIntPtr uFlag, out __MIDL___MIDL_itf_mshtml_0001_0042_0007 pDst, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,true,true);
			pDst = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0007();
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, hIMC, pSrc, uBufLen, uFlag, pDst, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetConversionListA", paramsArray, modifiers);
			pDst = (__MIDL___MIDL_itf_mshtml_0001_0042_0007)paramsArray[5];
			puCopied = (UIntPtr)paramsArray[6];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

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
		public virtual Int32 GetConversionListW(object hKL, Int32 hIMC, string pSrc, UIntPtr uBufLen, UIntPtr uFlag, out __MIDL___MIDL_itf_mshtml_0001_0042_0007 pDst, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,true,true);
			pDst = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0007();
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, hIMC, pSrc, uBufLen, uFlag, pDst, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetConversionListW", paramsArray, modifiers);
			pDst = (__MIDL___MIDL_itf_mshtml_0001_0042_0007)paramsArray[5];
			puCopied = (UIntPtr)paramsArray[6];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pfdwConversion">Int32 pfdwConversion</param>
		/// <param name="pfdwSentence">Int32 pfdwSentence</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetConversionStatus(Int32 hIMC, out Int32 pfdwConversion, out Int32 pfdwSentence)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true);
			pfdwConversion = 0;
			pfdwSentence = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, pfdwConversion, pfdwSentence);
			object returnItem = Invoker.MethodReturn(this, "GetConversionStatus", paramsArray, modifiers);
			pfdwConversion = (Int32)paramsArray[1];
			pfdwSentence = (Int32)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="phDefWnd">_RemotableHandle phDefWnd</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetDefaultIMEWnd(_RemotableHandle hWnd, out _RemotableHandle phDefWnd)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			phDefWnd = new NetOffice.MSHTMLApi._RemotableHandle();
			object[] paramsArray = Invoker.ValidateParamsArray(hWnd, phDefWnd);
			object returnItem = Invoker.MethodReturn(this, "GetDefaultIMEWnd", paramsArray, modifiers);
			phDefWnd = (_RemotableHandle)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="szDescription">string szDescription</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetDescriptionA(object hKL, UIntPtr uBufLen, out string szDescription, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true);
			szDescription = string.Empty;
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, uBufLen, szDescription, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetDescriptionA", paramsArray, modifiers);
			szDescription = paramsArray[2] as string;
			puCopied = (UIntPtr)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="szDescription">string szDescription</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetDescriptionW(object hKL, UIntPtr uBufLen, out string szDescription, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true);
			szDescription = string.Empty;
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, uBufLen, szDescription, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetDescriptionW", paramsArray, modifiers);
			szDescription = paramsArray[2] as string;
			puCopied = (UIntPtr)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="dwBufLen">Int32 dwBufLen</param>
		/// <param name="pBuf">string pBuf</param>
		/// <param name="pdwResult">Int32 pdwResult</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetGuideLineA(Int32 hIMC, Int32 dwIndex, Int32 dwBufLen, out string pBuf, out Int32 pdwResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true);
			pBuf = string.Empty;
			pdwResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwIndex, dwBufLen, pBuf, pdwResult);
			object returnItem = Invoker.MethodReturn(this, "GetGuideLineA", paramsArray, modifiers);
			pBuf = paramsArray[3] as string;
			pdwResult = (Int32)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="dwBufLen">Int32 dwBufLen</param>
		/// <param name="pBuf">string pBuf</param>
		/// <param name="pdwResult">Int32 pdwResult</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetGuideLineW(Int32 hIMC, Int32 dwIndex, Int32 dwBufLen, out string pBuf, out Int32 pdwResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true);
			pBuf = string.Empty;
			pdwResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwIndex, dwBufLen, pBuf, pdwResult);
			object returnItem = Invoker.MethodReturn(this, "GetGuideLineW", paramsArray, modifiers);
			pBuf = paramsArray[3] as string;
			pdwResult = (Int32)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="szFileName">string szFileName</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetIMEFileNameA(object hKL, UIntPtr uBufLen, out string szFileName, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true);
			szFileName = string.Empty;
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, uBufLen, szFileName, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetIMEFileNameA", paramsArray, modifiers);
			szFileName = paramsArray[2] as string;
			puCopied = (UIntPtr)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="uBufLen">UIntPtr uBufLen</param>
		/// <param name="szFileName">string szFileName</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetIMEFileNameW(object hKL, UIntPtr uBufLen, out string szFileName, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true);
			szFileName = string.Empty;
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, uBufLen, szFileName, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetIMEFileNameW", paramsArray);
			szFileName = paramsArray[2] as string;
			puCopied = (UIntPtr)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetOpenStatus(Int32 hIMC)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetOpenStatus", hIMC);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="fdwIndex">Int32 fdwIndex</param>
		/// <param name="pdwProperty">Int32 pdwProperty</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetProperty(object hKL, Int32 fdwIndex, out Int32 pdwProperty)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			pdwProperty = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, fdwIndex, pdwProperty);
			object returnItem = Invoker.MethodReturn(this, "GetProperty", paramsArray, modifiers);
			pdwProperty = (Int32)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="nItem">UIntPtr nItem</param>
		/// <param name="pStyleBuf">__MIDL___MIDL_itf_mshtml_0001_0042_0008 pStyleBuf</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetRegisterWordStyleA(object hKL, UIntPtr nItem, out __MIDL___MIDL_itf_mshtml_0001_0042_0008 pStyleBuf, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true);
			pStyleBuf = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0008();
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, nItem, pStyleBuf, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetRegisterWordStyleA", paramsArray, modifiers);
			pStyleBuf = (__MIDL___MIDL_itf_mshtml_0001_0042_0008)paramsArray[2];
			puCopied = (UIntPtr)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="nItem">UIntPtr nItem</param>
		/// <param name="pStyleBuf">__MIDL___MIDL_itf_mshtml_0001_0042_0009 pStyleBuf</param>
		/// <param name="puCopied">UIntPtr puCopied</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetRegisterWordStyleW(object hKL, UIntPtr nItem, out __MIDL___MIDL_itf_mshtml_0001_0042_0009 pStyleBuf, out UIntPtr puCopied)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true);
			pStyleBuf = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0009();
			puCopied = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, nItem, pStyleBuf, puCopied);
			object returnItem = Invoker.MethodReturn(this, "GetRegisterWordStyleW", paramsArray, modifiers);
			pStyleBuf = (__MIDL___MIDL_itf_mshtml_0001_0042_0009)paramsArray[2];
			puCopied = (UIntPtr)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pptPos">tagPOINT pptPos</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetStatusWindowPos(Int32 hIMC, out tagPOINT pptPos)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pptPos = new NetOffice.MSHTMLApi.tagPOINT();
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, pptPos);
			object returnItem = Invoker.MethodReturn(this, "GetStatusWindowPos", paramsArray, modifiers);
			pptPos = (tagPOINT)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="puVirtualKey">UIntPtr puVirtualKey</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetVirtualKey(_RemotableHandle hWnd, out UIntPtr puVirtualKey)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			puVirtualKey = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hWnd, puVirtualKey);
			object returnItem = Invoker.MethodReturn(this, "GetVirtualKey", paramsArray, modifiers);
			puVirtualKey = (UIntPtr)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="szIMEFileName">string szIMEFileName</param>
		/// <param name="szLayoutText">string szLayoutText</param>
		/// <param name="phKL">object phKL</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 InstallIMEA(string szIMEFileName, string szLayoutText, out object phKL)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			phKL = null;
			object[] paramsArray = Invoker.ValidateParamsArray(szIMEFileName, szLayoutText, phKL);
			object returnItem = Invoker.MethodReturn(this, "InstallIMEA", paramsArray, modifiers);
			phKL = (object)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="szIMEFileName">string szIMEFileName</param>
		/// <param name="szLayoutText">string szLayoutText</param>
		/// <param name="phKL">object phKL</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 InstallIMEW(string szIMEFileName, string szLayoutText, out object phKL)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			phKL = null;
			object[] paramsArray = Invoker.ValidateParamsArray(szIMEFileName, szLayoutText, phKL);
			object returnItem = Invoker.MethodReturn(this, "InstallIMEW", paramsArray, modifiers);
			phKL = (object)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsIME(object hKL)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "IsIME", hKL);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWndIME">_RemotableHandle hWndIME</param>
		/// <param name="msg">UIntPtr msg</param>
		/// <param name="wParam">Int32 wParam</param>
		/// <param name="lParam">Int32 lParam</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsUIMessageA(_RemotableHandle hWndIME, UIntPtr msg, Int32 wParam, Int32 lParam)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "IsUIMessageA", hWndIME, msg, wParam, lParam);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWndIME">_RemotableHandle hWndIME</param>
		/// <param name="msg">UIntPtr msg</param>
		/// <param name="wParam">Int32 wParam</param>
		/// <param name="lParam">Int32 lParam</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsUIMessageW(_RemotableHandle hWndIME, UIntPtr msg, Int32 wParam, Int32 lParam)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "IsUIMessageW", hWndIME, msg, wParam, lParam);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwAction">Int32 dwAction</param>
		/// <param name="dwIndex">Int32 dwIndex</param>
		/// <param name="dwValue">Int32 dwValue</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 NotifyIME(Int32 hIMC, Int32 dwAction, Int32 dwIndex, Int32 dwValue)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "NotifyIME", hIMC, dwAction, dwIndex, dwValue);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="szReading">string szReading</param>
		/// <param name="dwStyle">Int32 dwStyle</param>
		/// <param name="szRegister">string szRegister</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 RegisterWordA(object hKL, string szReading, Int32 dwStyle, string szRegister)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RegisterWordA", hKL, szReading, dwStyle, szRegister);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="szReading">string szReading</param>
		/// <param name="dwStyle">Int32 dwStyle</param>
		/// <param name="szRegister">string szRegister</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 RegisterWordW(object hKL, string szReading, Int32 dwStyle, string szRegister)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RegisterWordW", hKL, szReading, dwStyle, szRegister);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="hIMC">Int32 hIMC</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 ReleaseContext(_RemotableHandle hWnd, Int32 hIMC)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ReleaseContext", hWnd, hIMC);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pCandidate">__MIDL___MIDL_itf_mshtml_0001_0042_0005 pCandidate</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SetCandidateWindow(Int32 hIMC, __MIDL___MIDL_itf_mshtml_0001_0042_0005 pCandidate)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetCandidateWindow", hIMC, pCandidate);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="plf">__MIDL___MIDL_itf_mshtml_0001_0042_0003 plf</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SetCompositionFontA(Int32 hIMC, __MIDL___MIDL_itf_mshtml_0001_0042_0003 plf)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetCompositionFontA", hIMC, plf);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="plf">__MIDL___MIDL_itf_mshtml_0001_0042_0004 plf</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SetCompositionFontW(Int32 hIMC, __MIDL___MIDL_itf_mshtml_0001_0042_0004 plf)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetCompositionFontW", hIMC, plf);
		}

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
		public virtual Int32 SetCompositionStringA(Int32 hIMC, Int32 dwIndex, object pComp, Int32 dwCompLen, object pRead, Int32 dwReadLen)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetCompositionStringA", new object[]{ hIMC, dwIndex, pComp, dwCompLen, pRead, dwReadLen });
		}

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
		public virtual Int32 SetCompositionStringW(Int32 hIMC, Int32 dwIndex, object pComp, Int32 dwCompLen, object pRead, Int32 dwReadLen)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetCompositionStringW", new object[]{ hIMC, dwIndex, pComp, dwCompLen, pRead, dwReadLen });
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pCompForm">__MIDL___MIDL_itf_mshtml_0001_0042_0006 pCompForm</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SetCompositionWindow(Int32 hIMC, __MIDL___MIDL_itf_mshtml_0001_0042_0006 pCompForm)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetCompositionWindow", hIMC, pCompForm);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="fdwConversion">Int32 fdwConversion</param>
		/// <param name="fdwSentence">Int32 fdwSentence</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SetConversionStatus(Int32 hIMC, Int32 fdwConversion, Int32 fdwSentence)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetConversionStatus", hIMC, fdwConversion, fdwSentence);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="fOpen">Int32 fOpen</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SetOpenStatus(Int32 hIMC, Int32 fOpen)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetOpenStatus", hIMC, fOpen);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="pptPos">tagPOINT pptPos</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SetStatusWindowPos(Int32 hIMC, tagPOINT pptPos)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetStatusWindowPos", hIMC, pptPos);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="dwHotKeyID">Int32 dwHotKeyID</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SimulateHotKey(_RemotableHandle hWnd, Int32 dwHotKeyID)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SimulateHotKey", hWnd, dwHotKeyID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="szReading">string szReading</param>
		/// <param name="dwStyle">Int32 dwStyle</param>
		/// <param name="szUnregister">string szUnregister</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 UnregisterWordA(object hKL, string szReading, Int32 dwStyle, string szUnregister)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "UnregisterWordA", hKL, szReading, dwStyle, szUnregister);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="szReading">string szReading</param>
		/// <param name="dwStyle">Int32 dwStyle</param>
		/// <param name="szUnregister">string szUnregister</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 UnregisterWordW(object hKL, string szReading, Int32 dwStyle, string szUnregister)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "UnregisterWordW", hKL, szReading, dwStyle, szUnregister);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fRestoreLayout">Int32 fRestoreLayout</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 Activate(Int32 fRestoreLayout)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Activate", fRestoreLayout);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 Deactivate()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Deactivate");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="msg">UIntPtr msg</param>
		/// <param name="wParam">Int32 wParam</param>
		/// <param name="lParam">Int32 lParam</param>
		/// <param name="plResult">Int32 plResult</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 OnDefWindowProc(_RemotableHandle hWnd, UIntPtr msg, Int32 wParam, Int32 lParam, out Int32 plResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true);
			plResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hWnd, msg, wParam, lParam, plResult);
			object returnItem = Invoker.MethodReturn(this, "OnDefWindowProc", paramsArray, modifiers);
			plResult = (Int32)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="aaClassList">Int16 aaClassList</param>
		/// <param name="uSize">UIntPtr uSize</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 FilterClientWindows(Int16 aaClassList, UIntPtr uSize)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "FilterClientWindows", aaClassList, uSize);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="uCodePage">UIntPtr uCodePage</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetCodePageA(object hKL, out UIntPtr uCodePage)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			uCodePage = UIntPtr.Zero;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, uCodePage);
			object returnItem = Invoker.MethodReturn(this, "GetCodePageA", paramsArray, modifiers);
			uCodePage = (UIntPtr)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hKL">object hKL</param>
		/// <param name="plid">Int16 plid</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetLangId(object hKL, out Int16 plid)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			plid = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hKL, plid);
			object returnItem = Invoker.MethodReturn(this, "GetLangId", paramsArray, modifiers);
			plid = (Int16)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hWnd">_RemotableHandle hWnd</param>
		/// <param name="hIMC">Int32 hIMC</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 AssociateContextEx(_RemotableHandle hWnd, Int32 hIMC, Int32 dwFlags)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "AssociateContextEx", hWnd, hIMC, dwFlags);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="idThread">Int32 idThread</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 DisableIME(Int32 idThread)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DisableIME", idThread);
		}

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
		public virtual Int32 GetImeMenuItemsA(Int32 hIMC, Int32 dwFlags, Int32 dwType, __MIDL___MIDL_itf_mshtml_0001_0042_0010 pImeParentMenu, out __MIDL___MIDL_itf_mshtml_0001_0042_0010 pImeMenu, Int32 dwSize, out Int32 pdwResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true,false,true);
			pImeMenu = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0010();
			pdwResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwFlags, dwType, pImeParentMenu, pImeMenu, dwSize, pdwResult);
			object returnItem = Invoker.MethodReturn(this, "GetImeMenuItemsA", paramsArray, modifiers);
			pImeMenu = (__MIDL___MIDL_itf_mshtml_0001_0042_0010)paramsArray[4];
			pdwResult = (Int32)paramsArray[6];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

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
		public virtual Int32 GetImeMenuItemsW(Int32 hIMC, Int32 dwFlags, Int32 dwType, __MIDL___MIDL_itf_mshtml_0001_0042_0011 pImeParentMenu, out __MIDL___MIDL_itf_mshtml_0001_0042_0011 pImeMenu, Int32 dwSize, out Int32 pdwResult)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true,false,true);
			pImeMenu = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0011();
			pdwResult = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(hIMC, dwFlags, dwType, pImeParentMenu, pImeMenu, dwSize, pdwResult);
			object returnItem = Invoker.MethodReturn(this, "GetImeMenuItemsW", paramsArray, modifiers);
			pImeMenu = (__MIDL___MIDL_itf_mshtml_0001_0042_0011)paramsArray[4];
			pdwResult = (Int32)paramsArray[6];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="idThread">Int32 idThread</param>
		/// <param name="ppEnum">NetOffice.MSHTMLApi.IEnumInputContext ppEnum</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 EnumInputContext(Int32 idThread, out NetOffice.MSHTMLApi.IEnumInputContext ppEnum)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			ppEnum = null;
			object[] paramsArray = Invoker.ValidateParamsArray(idThread, ppEnum);
			object returnItem = Invoker.MethodReturn(this, "EnumInputContext", paramsArray, modifiers);
			ppEnum = (NetOffice.MSHTMLApi.IEnumInputContext)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

