using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.PowerPointApi
{
	///<summary>
	/// DispatchInterface ProtectedViewWindows 
	/// SupportByVersion PowerPoint, 14
	///</summary>
	[SupportByVersionAttribute("PowerPoint", 14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class ProtectedViewWindows : Collection
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
                    _type = typeof(ProtectedViewWindows);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProtectedViewWindows(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProtectedViewWindows(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProtectedViewWindows(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProtectedViewWindows() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProtectedViewWindows(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.PowerPointApi.Application newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Application.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("PowerPoint", 14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.PowerPointApi.ProtectedViewWindow this[Int32 index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.PowerPointApi.ProtectedViewWindow newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.ProtectedViewWindow.LateBindingApiWrapperType) as NetOffice.PowerPointApi.ProtectedViewWindow;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="readPassword">optional string ReadPassword = </param>
		/// <param name="openAndRepair">optional NetOffice.OfficeApi.Enums.MsoTriState OpenAndRepair = 0</param>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.ProtectedViewWindow Open(string fileName, string readPassword, NetOffice.OfficeApi.Enums.MsoTriState openAndRepair)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, readPassword, openAndRepair);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.PowerPointApi.ProtectedViewWindow newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.ProtectedViewWindow.LateBindingApiWrapperType) as NetOffice.PowerPointApi.ProtectedViewWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.ProtectedViewWindow Open(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.PowerPointApi.ProtectedViewWindow newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.ProtectedViewWindow.LateBindingApiWrapperType) as NetOffice.PowerPointApi.ProtectedViewWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="readPassword">optional string ReadPassword = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.ProtectedViewWindow Open(string fileName, string readPassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, readPassword);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.PowerPointApi.ProtectedViewWindow newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.ProtectedViewWindow.LateBindingApiWrapperType) as NetOffice.PowerPointApi.ProtectedViewWindow;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}