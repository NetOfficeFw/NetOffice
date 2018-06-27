using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface ISelectionServicesListener 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class ISelectionServicesListener : COMObject, NetOffice.MSHTMLApi.ISelectionServicesListener
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
                    _contractType = typeof(NetOffice.MSHTMLApi.ISelectionServicesListener);
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
                    _type = typeof(ISelectionServicesListener);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ISelectionServicesListener() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 BeginSelectionUndo()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "BeginSelectionUndo");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 EndSelectionUndo()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "EndSelectionUndo");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIElementStart">NetOffice.MSHTMLApi.IMarkupPointer pIElementStart</param>
		/// <param name="pIElementEnd">NetOffice.MSHTMLApi.IMarkupPointer pIElementEnd</param>
		/// <param name="pIElementContentStart">NetOffice.MSHTMLApi.IMarkupPointer pIElementContentStart</param>
		/// <param name="pIElementContentEnd">NetOffice.MSHTMLApi.IMarkupPointer pIElementContentEnd</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 OnSelectedElementExit(NetOffice.MSHTMLApi.IMarkupPointer pIElementStart, NetOffice.MSHTMLApi.IMarkupPointer pIElementEnd, NetOffice.MSHTMLApi.IMarkupPointer pIElementContentStart, NetOffice.MSHTMLApi.IMarkupPointer pIElementContentEnd)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "OnSelectedElementExit", pIElementStart, pIElementEnd, pIElementContentStart, pIElementContentEnd);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="eType">NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType</param>
		/// <param name="pIListener">NetOffice.MSHTMLApi.ISelectionServicesListener pIListener</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 OnChangeType(NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType, NetOffice.MSHTMLApi.ISelectionServicesListener pIListener)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "OnChangeType", eType, pIListener);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pTypeDetail">string pTypeDetail</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetTypeDetail(out string pTypeDetail)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pTypeDetail = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(pTypeDetail);
			object returnItem = Invoker.MethodReturn(this, "GetTypeDetail", paramsArray, modifiers);
			pTypeDetail = paramsArray[0] as string;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

