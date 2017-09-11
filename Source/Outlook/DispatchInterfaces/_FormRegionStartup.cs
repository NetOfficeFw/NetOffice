using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _FormRegionStartup 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _FormRegionStartup : COMObject
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
                    _type = typeof(_FormRegionStartup);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _FormRegionStartup(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866042.aspx </remarks>
		/// <param name="formRegionName">string formRegionName</param>
		/// <param name="item">object item</param>
		/// <param name="lCID">Int32 lCID</param>
		/// <param name="formRegionMode">NetOffice.OutlookApi.Enums.OlFormRegionMode formRegionMode</param>
		/// <param name="formRegionSize">NetOffice.OutlookApi.Enums.OlFormRegionSize formRegionSize</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public object GetFormRegionStorage(string formRegionName, object item, Int32 lCID, NetOffice.OutlookApi.Enums.OlFormRegionMode formRegionMode, NetOffice.OutlookApi.Enums.OlFormRegionSize formRegionSize)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetFormRegionStorage", new object[]{ formRegionName, item, lCID, formRegionMode, formRegionSize });
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869072.aspx </remarks>
		/// <param name="formRegion">NetOffice.OutlookApi.FormRegion formRegion</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void BeforeFormRegionShow(NetOffice.OutlookApi.FormRegion formRegion)
		{
			 Factory.ExecuteMethod(this, "BeforeFormRegionShow", formRegion);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869502.aspx </remarks>
		/// <param name="formRegionName">string formRegionName</param>
		/// <param name="lCID">Int32 lCID</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public object GetFormRegionManifest(string formRegionName, Int32 lCID)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetFormRegionManifest", formRegionName, lCID);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868914.aspx </remarks>
		/// <param name="formRegionName">string formRegionName</param>
		/// <param name="lCID">Int32 lCID</param>
		/// <param name="icon">NetOffice.OutlookApi.Enums.OlFormRegionIcon icon</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public object GetFormRegionIcon(string formRegionName, Int32 lCID, NetOffice.OutlookApi.Enums.OlFormRegionIcon icon)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetFormRegionIcon", formRegionName, lCID, icon);
		}

		#endregion

		#pragma warning restore
	}
}
