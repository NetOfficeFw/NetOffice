using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// Interface IDocumentInspector 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861808.aspx </remarks>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class IDocumentInspector : COMObject
	{
		#pragma warning disable

		#region Type Information

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
                    _type = typeof(IDocumentInspector);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IDocumentInspector(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IDocumentInspector(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDocumentInspector(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDocumentInspector(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDocumentInspector(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDocumentInspector(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDocumentInspector() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDocumentInspector(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862465.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="desc">string desc</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 GetInfo(out string name, out string desc)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true);
			name = string.Empty;
			desc = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(name, desc);
			object returnItem = Invoker.MethodReturn(this, "GetInfo", paramsArray, modifiers);
			name = paramsArray[0] as string;
			desc = paramsArray[1] as string;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861133.aspx </remarks>
		/// <param name="doc">object doc</param>
		/// <param name="status">NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status</param>
		/// <param name="result">string result</param>
		/// <param name="action">string action</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 Inspect(object doc, out NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status, out string result, out string action)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true,true);
			status = 0;
			result = string.Empty;
			action = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(doc, status, result, action);
			object returnItem = Invoker.MethodReturn(this, "Inspect", paramsArray, modifiers);
			status = (NetOffice.OfficeApi.Enums.MsoDocInspectorStatus)paramsArray[1];
			result = paramsArray[2] as string;
			action = paramsArray[3] as string;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864114.aspx </remarks>
		/// <param name="doc">object doc</param>
		/// <param name="hwnd">Int32 hwnd</param>
		/// <param name="status">NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status</param>
		/// <param name="result">string result</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 Fix(object doc, Int32 hwnd, out NetOffice.OfficeApi.Enums.MsoDocInspectorStatus status, out string result)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true);
			status = 0;
			result = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(doc, hwnd, status, result);
			object returnItem = Invoker.MethodReturn(this, "Fix", paramsArray, modifiers);
			status = (NetOffice.OfficeApi.Enums.MsoDocInspectorStatus)paramsArray[2];
			result = paramsArray[3] as string;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}
