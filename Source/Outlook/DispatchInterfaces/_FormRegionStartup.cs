using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OutlookApi
{
	///<summary>
	/// DispatchInterface _FormRegionStartup 
	/// SupportByVersion Outlook, 12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Outlook", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _FormRegionStartup : COMObject
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
                    _type = typeof(_FormRegionStartup);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _FormRegionStartup(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _FormRegionStartup(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _FormRegionStartup(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _FormRegionStartup(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _FormRegionStartup(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _FormRegionStartup() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _FormRegionStartup(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff866042.aspx
		/// </summary>
		/// <param name="formRegionName">string FormRegionName</param>
		/// <param name="item">object Item</param>
		/// <param name="lCID">Int32 LCID</param>
		/// <param name="formRegionMode">NetOffice.OutlookApi.Enums.OlFormRegionMode FormRegionMode</param>
		/// <param name="formRegionSize">NetOffice.OutlookApi.Enums.OlFormRegionSize FormRegionSize</param>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public object GetFormRegionStorage(string formRegionName, object item, Int32 lCID, NetOffice.OutlookApi.Enums.OlFormRegionMode formRegionMode, NetOffice.OutlookApi.Enums.OlFormRegionSize formRegionSize)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formRegionName, item, lCID, formRegionMode, formRegionSize);
			object returnItem = Invoker.MethodReturn(this, "GetFormRegionStorage", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869072.aspx
		/// </summary>
		/// <param name="formRegion">NetOffice.OutlookApi.FormRegion FormRegion</param>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public void BeforeFormRegionShow(NetOffice.OutlookApi.FormRegion formRegion)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formRegion);
			Invoker.Method(this, "BeforeFormRegionShow", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869502.aspx
		/// </summary>
		/// <param name="formRegionName">string FormRegionName</param>
		/// <param name="lCID">Int32 LCID</param>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public object GetFormRegionManifest(string formRegionName, Int32 lCID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formRegionName, lCID);
			object returnItem = Invoker.MethodReturn(this, "GetFormRegionManifest", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff868914.aspx
		/// </summary>
		/// <param name="formRegionName">string FormRegionName</param>
		/// <param name="lCID">Int32 LCID</param>
		/// <param name="icon">NetOffice.OutlookApi.Enums.OlFormRegionIcon Icon</param>
		[SupportByVersionAttribute("Outlook", 12,14,15,16)]
		public object GetFormRegionIcon(string formRegionName, Int32 lCID, NetOffice.OutlookApi.Enums.OlFormRegionIcon icon)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formRegionName, lCID, icon);
			object returnItem = Invoker.MethodReturn(this, "GetFormRegionIcon", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		#endregion
		#pragma warning restore
	}
}