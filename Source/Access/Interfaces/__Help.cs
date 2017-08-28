using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// Interface __Help 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class __Help : COMObject
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
                    _type = typeof(__Help);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public __Help(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public __Help(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public __Help(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public __Help(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public __Help(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public __Help(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public __Help() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public __Help(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object FieldName
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FieldName");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "FieldName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object DataType
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DataType");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DataType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object Description
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Description");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Description", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object FieldSize
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FieldSize");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "FieldSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object NewValues
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "NewValues");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "NewValues", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object Required
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Required");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Required", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object AllowZeroLength
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "AllowZeroLength");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "AllowZeroLength", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object Indexed
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Indexed");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Indexed", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object DisplayControl
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DisplayControl");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DisplayControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object ReplicationConflictFunction
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ReplicationConflictFunction");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ReplicationConflictFunction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object ProjectName
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ProjectName");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ProjectName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object MDE
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "MDE");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "MDE", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object AllowBreakIntoCode
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "AllowBreakIntoCode");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "AllowBreakIntoCode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object AllowBuiltInToolbars
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "AllowBuiltInToolbars");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "AllowBuiltInToolbars", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object AllowBypassKey
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "AllowBypassKey");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "AllowBypassKey", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object AllowFullMenus
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "AllowFullMenus");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "AllowFullMenus", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object AllowSpecialKeys
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "AllowSpecialKeys");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "AllowSpecialKeys", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object AllowToolbarChanges
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "AllowToolbarChanges");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "AllowToolbarChanges", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object StartUpForm
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "StartUpForm");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "StartUpForm", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object StartUpMenuBar
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "StartUpMenuBar");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "StartUpMenuBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object StartUpShortcutMenuBar
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "StartUpShortcutMenuBar");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "StartUpShortcutMenuBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object StartUpShowDBWindow
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "StartUpShowDBWindow");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "StartUpShowDBWindow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object StartUpShowStatusBar
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "StartUpShowStatusBar");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "StartUpShowStatusBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object AppIcon
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "AppIcon");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "AppIcon", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object AppTitle
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "AppTitle");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "AppTitle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object ODBCConnectStr
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ODBCConnectStr");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ODBCConnectStr", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object LogMessages
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LogMessages");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "LogMessages", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
