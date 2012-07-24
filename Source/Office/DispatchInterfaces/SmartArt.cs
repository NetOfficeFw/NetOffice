using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// DispatchInterface SmartArt 
	/// SupportByVersion Office, 14,15
	///</summary>
	[SupportByVersionAttribute("Office", 14,15)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class SmartArt : _IMsoDispObj
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
                    _type = typeof(SmartArt);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SmartArt(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SmartArt(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SmartArt(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SmartArt() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SmartArt(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 14, 15
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15)]
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

		/// <summary>
		/// SupportByVersion Office 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15)]
		public NetOffice.OfficeApi.SmartArtNodes AllNodes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AllNodes", paramsArray);
				NetOffice.OfficeApi.SmartArtNodes newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SmartArtNodes.LateBindingApiWrapperType) as NetOffice.OfficeApi.SmartArtNodes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15)]
		public NetOffice.OfficeApi.SmartArtNodes Nodes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Nodes", paramsArray);
				NetOffice.OfficeApi.SmartArtNodes newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SmartArtNodes.LateBindingApiWrapperType) as NetOffice.OfficeApi.SmartArtNodes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15)]
		public NetOffice.OfficeApi.SmartArtLayout Layout
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Layout", paramsArray);
				NetOffice.OfficeApi.SmartArtLayout newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SmartArtLayout.LateBindingApiWrapperType) as NetOffice.OfficeApi.SmartArtLayout;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Layout", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15)]
		public NetOffice.OfficeApi.SmartArtQuickStyle QuickStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "QuickStyle", paramsArray);
				NetOffice.OfficeApi.SmartArtQuickStyle newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SmartArtQuickStyle.LateBindingApiWrapperType) as NetOffice.OfficeApi.SmartArtQuickStyle;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "QuickStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15)]
		public NetOffice.OfficeApi.SmartArtColor Color
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Color", paramsArray);
				NetOffice.OfficeApi.SmartArtColor newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.SmartArtColor.LateBindingApiWrapperType) as NetOffice.OfficeApi.SmartArtColor;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Color", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15)]
		public NetOffice.OfficeApi.Enums.MsoTriState Reverse
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Reverse", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoTriState)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Reverse", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 14, 15
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15)]
		public void Reset()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Reset", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}