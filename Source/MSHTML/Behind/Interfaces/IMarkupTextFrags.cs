using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IMarkupTextFrags 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IMarkupTextFrags : COMObject, NetOffice.MSHTMLApi.IMarkupTextFrags
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IMarkupTextFrags);
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
                    _type = typeof(IMarkupTextFrags);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IMarkupTextFrags() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pcFrags">Int32 pcFrags</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetTextFragCount(out Int32 pcFrags)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pcFrags = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pcFrags);
			object returnItem = Invoker.MethodReturn(this, "GetTextFragCount", paramsArray, modifiers);
			pcFrags = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="iFrag">Int32 iFrag</param>
		/// <param name="pbstrFrag">string pbstrFrag</param>
		/// <param name="pPointerFrag">NetOffice.MSHTMLApi.IMarkupPointer pPointerFrag</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetTextFrag(Int32 iFrag, out string pbstrFrag, NetOffice.MSHTMLApi.IMarkupPointer pPointerFrag)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,false);
			pbstrFrag = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(iFrag, pbstrFrag, pPointerFrag);
			object returnItem = Invoker.MethodReturn(this, "GetTextFrag", paramsArray, modifiers);
			pbstrFrag = paramsArray[1] as string;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="iFrag">Int32 iFrag</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 RemoveTextFrag(Int32 iFrag)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RemoveTextFrag", iFrag);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="iFrag">Int32 iFrag</param>
		/// <param name="bstrInsert">string bstrInsert</param>
		/// <param name="pPointerInsert">NetOffice.MSHTMLApi.IMarkupPointer pPointerInsert</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 InsertTextFrag(Int32 iFrag, string bstrInsert, NetOffice.MSHTMLApi.IMarkupPointer pPointerInsert)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "InsertTextFrag", iFrag, bstrInsert, pPointerInsert);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerFind">NetOffice.MSHTMLApi.IMarkupPointer pPointerFind</param>
		/// <param name="piFrag">Int32 piFrag</param>
		/// <param name="pfFragFound">Int32 pfFragFound</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 FindTextFragFromMarkupPointer(NetOffice.MSHTMLApi.IMarkupPointer pPointerFind, out Int32 piFrag, out Int32 pfFragFound)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true);
			piFrag = 0;
			pfFragFound = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pPointerFind, piFrag, pfFragFound);
			object returnItem = Invoker.MethodReturn(this, "FindTextFragFromMarkupPointer", paramsArray, modifiers);
			piFrag = (Int32)paramsArray[1];
			pfFragFound = (Int32)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

