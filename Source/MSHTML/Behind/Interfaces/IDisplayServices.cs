using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IDisplayServices 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IDisplayServices : COMObject, NetOffice.MSHTMLApi.IDisplayServices
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IDisplayServices);
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
                    _type = typeof(IDisplayServices);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IDisplayServices() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppDispPointer">NetOffice.MSHTMLApi.IDisplayPointer ppDispPointer</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 CreateDisplayPointer(out NetOffice.MSHTMLApi.IDisplayPointer ppDispPointer)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppDispPointer = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppDispPointer);
			object returnItem = Invoker.MethodReturn(this, "CreateDisplayPointer", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppDispPointer = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IDisplayPointer>(this, paramsArray[0], typeof(NetOffice.MSHTMLApi.IDisplayPointer));
            else
                ppDispPointer = null;
            return NetRuntimeSystem.Convert.ToInt32(returnItem);
        }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pRect">tagRECT pRect</param>
		/// <param name="eSource">NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eSource</param>
		/// <param name="eDestination">NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eDestination</param>
		/// <param name="pIElement">NetOffice.MSHTMLApi.IHTMLElement pIElement</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 TransformRect(tagRECT pRect, NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eSource, NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eDestination, NetOffice.MSHTMLApi.IHTMLElement pIElement)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "TransformRect", pRect, eSource, eDestination, pIElement);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPoint">tagPOINT pPoint</param>
		/// <param name="eSource">NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eSource</param>
		/// <param name="eDestination">NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eDestination</param>
		/// <param name="pIElement">NetOffice.MSHTMLApi.IHTMLElement pIElement</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 TransformPoint(tagPOINT pPoint, NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eSource, NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eDestination, NetOffice.MSHTMLApi.IHTMLElement pIElement)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "TransformPoint", pPoint, eSource, eDestination, pIElement);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppCaret">NetOffice.MSHTMLApi.IHTMLCaret ppCaret</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetCaret(out NetOffice.MSHTMLApi.IHTMLCaret ppCaret)
		{
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			object[] paramsArray = Invoker.ValidateParamsArray(new object());
			object returnItem = Invoker.MethodReturn(this, "GetCaret", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppCaret = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IHTMLCaret>(this, paramsArray[0], typeof(NetOffice.MSHTMLApi.IHTMLCaret));
            else
                ppCaret = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);

            
        }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointer">NetOffice.MSHTMLApi.IMarkupPointer pPointer</param>
		/// <param name="ppComputedStyle">NetOffice.MSHTMLApi.IHTMLComputedStyle ppComputedStyle</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetComputedStyle(NetOffice.MSHTMLApi.IMarkupPointer pPointer, out NetOffice.MSHTMLApi.IHTMLComputedStyle ppComputedStyle)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			object[] paramsArray = Invoker.ValidateParamsArray(pPointer, new object());
			object returnItem = Invoker.MethodReturn(this, "GetComputedStyle", paramsArray, modifiers);
            if (paramsArray[1] is MarshalByRefObject)
                ppComputedStyle = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IHTMLComputedStyle>(this, paramsArray[1], typeof(NetOffice.MSHTMLApi.IHTMLComputedStyle));
            else
                ppComputedStyle = null;
            return NetRuntimeSystem.Convert.ToInt32(returnItem);
        }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIElement">NetOffice.MSHTMLApi.IHTMLElement pIElement</param>
		/// <param name="rect">tagRECT rect</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 ScrollRectIntoView(NetOffice.MSHTMLApi.IHTMLElement pIElement, tagRECT rect)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ScrollRectIntoView", pIElement, rect);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIElement">NetOffice.MSHTMLApi.IHTMLElement pIElement</param>
		/// <param name="pfHasFlowLayout">Int32 pfHasFlowLayout</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 HasFlowLayout(NetOffice.MSHTMLApi.IHTMLElement pIElement, out Int32 pfHasFlowLayout)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pfHasFlowLayout = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pIElement, pfHasFlowLayout);
			object returnItem = Invoker.MethodReturn(this, "HasFlowLayout", paramsArray, modifiers);
			pfHasFlowLayout = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

