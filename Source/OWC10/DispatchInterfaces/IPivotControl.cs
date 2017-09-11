using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface IPivotControl 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IPivotControl : COMObject
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
                    _type = typeof(IPivotControl);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IPivotControl(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IPivotControl(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotControl(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotControl(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotControl(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotControl(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotControl() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotControl(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotView ActiveView
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotView>(this, "ActiveView", NetOffice.OWC10Api.PivotView.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		public object Selection
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Selection");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Selection", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string DataMember
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataMember");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataMember", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotData ActiveData
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotData>(this, "ActiveData", NetOffice.OWC10Api.PivotData.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string Version
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool HasDetails
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasDetails");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayToolbar
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayToolbar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayToolbar", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowGrouping
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowGrouping");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowGrouping", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowFiltering
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowFiltering");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowFiltering", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowDetails
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowDetails");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowDetails", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowPropertyToolbox
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowPropertyToolbox");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowPropertyToolbox", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowCustomOrdering
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowCustomOrdering");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowCustomOrdering", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AutoFit
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFit", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.MSDATASRCApi.DataSource DataSource
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSDATASRCApi.DataSource>(this, "DataSource", NetOffice.MSDATASRCApi.DataSource.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "DataSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object BackColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BackColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayExpandIndicator
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayExpandIndicator");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayExpandIndicator", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool RightToLeft
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RightToLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RightToLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 MaxWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MaxWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaxWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 MaxHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MaxHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaxHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Width
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Width");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Height
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Height");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Height", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string XMLData
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "XMLData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "XMLData", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayPropertyToolbox
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayPropertyToolbox");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayPropertyToolbox", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayFieldList
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayFieldList");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayFieldList", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Constants
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Constants");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 MajorVersion
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MajorVersion");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string MinorVersion
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MinorVersion");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string BuildNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BuildNumber");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string ConnectionString
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ConnectionString");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConnectionString", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string CommandText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CommandText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandText", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ProviderType ProviderType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ProviderType>(this, "ProviderType");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.Enums.PivotTableMemberExpandEnum MemberExpand
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotTableMemberExpandEnum>(this, "MemberExpand");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MemberExpand", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ADODBApi.Connection Connection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Connection>(this, "Connection", NetOffice.ADODBApi.Connection.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Connection", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string RevisionNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RevisionNumber");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayAlerts
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayAlerts");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayAlerts", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object DataMemberStrings
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DataMemberStrings");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.PivotClassFactory ClassFactory
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotClassFactory>(this, "ClassFactory", NetOffice.OWC10Api.PivotClassFactory.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "ClassFactory", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Left
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Left");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Top
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Top");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Hwnd
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Hwnd");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		public object ActiveObject
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ActiveObject");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "ActiveObject", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.OCCommands Commands
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.OCCommands>(this, "Commands", NetOffice.OWC10Api.OCCommands.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool UserMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UserMode");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string DataMemberCaption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataMemberCaption");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataMemberCaption", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object DataSourceEx
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "DataSourceEx");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "DataSourceEx", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool IsDirty
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsDirty");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IsDirty", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string CubeProvider
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CubeProvider");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CubeProvider", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string SelectionType
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SelectionType");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayScreenTips
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayScreenTips");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayScreenTips", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ViewOnlyMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ViewOnlyMode");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayDesignTimeUI
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayDesignTimeUI");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayDesignTimeUI", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.MSComctlLibApi.IToolbar Toolbar
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSComctlLibApi.IToolbar>(this, "Toolbar");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotEditModeEnum EditMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotEditModeEnum>(this, "EditMode");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string HTMLData
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HTMLData");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string DataSourceName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataSourceName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataSourceName", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool DisplayBranding
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayBranding");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayBranding", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayOfficeLogo
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayOfficeLogo");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayOfficeLogo", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = </param>
		/// <param name="action">optional NetOffice.OWC10Api.Enums.PivotExportActionEnum Action = 1</param>
		[SupportByVersion("OWC10", 1)]
		public void Export(object filename, object action)
		{
			 Factory.ExecuteMethod(this, "Export", filename, action);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Export()
		{
			 Factory.ExecuteMethod(this, "Export");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = </param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Export(object filename)
		{
			 Factory.ExecuteMethod(this, "Export", filename);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Refresh()
		{
			 Factory.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = pivot.gif</param>
		/// <param name="filterName">optional string FilterName = gif</param>
		/// <param name="width">optional Int32 Width = 1024</param>
		/// <param name="height">optional Int32 Height = 1024</param>
		[SupportByVersion("OWC10", 1)]
		public void ExportPicture(object filename, object filterName, object width, object height)
		{
			 Factory.ExecuteMethod(this, "ExportPicture", filename, filterName, width, height);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void ExportPicture()
		{
			 Factory.ExecuteMethod(this, "ExportPicture");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = pivot.gif</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void ExportPicture(object filename)
		{
			 Factory.ExecuteMethod(this, "ExportPicture", filename);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = pivot.gif</param>
		/// <param name="filterName">optional string FilterName = gif</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void ExportPicture(object filename, object filterName)
		{
			 Factory.ExecuteMethod(this, "ExportPicture", filename, filterName);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = pivot.gif</param>
		/// <param name="filterName">optional string FilterName = gif</param>
		/// <param name="width">optional Int32 Width = 1024</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void ExportPicture(object filename, object filterName, object width)
		{
			 Factory.ExecuteMethod(this, "ExportPicture", filename, filterName, width);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void LocateDataSource()
		{
			 Factory.ExecuteMethod(this, "LocateDataSource");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="selection">optional object Selection = null (Nothing in visual basic)</param>
		[SupportByVersion("OWC10", 1)]
		public void Copy(object selection)
		{
			 Factory.ExecuteMethod(this, "Copy", selection);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">NetOffice.OWC10Api.DropSource source</param>
		/// <param name="dragItem">object dragItem</param>
		/// <param name="target">NetOffice.OWC10Api.DropTarget target</param>
		/// <param name="dwLegalEffect">Int32 dwLegalEffect</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void DoDragDrop(NetOffice.OWC10Api.DropSource source, object dragItem, NetOffice.OWC10Api.DropTarget target, Int32 dwLegalEffect)
		{
			 Factory.ExecuteMethod(this, "DoDragDrop", source, dragItem, target, dwLegalEffect);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="selection">object selection</param>
		/// <param name="activeObject">object activeObject</param>
		/// <param name="scrollType">optional NetOffice.OWC10Api.Enums.PivotScrollTypeEnum ScrollType = 0</param>
		/// <param name="update">optional bool Update = true</param>
		/// <param name="notify">optional bool Notify = true</param>
		[SupportByVersion("OWC10", 1)]
		public void Select(object selection, object activeObject, object scrollType, object update, object notify)
		{
			 Factory.ExecuteMethod(this, "Select", new object[]{ selection, activeObject, scrollType, update, notify });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="selection">object selection</param>
		/// <param name="activeObject">object activeObject</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Select(object selection, object activeObject)
		{
			 Factory.ExecuteMethod(this, "Select", selection, activeObject);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="selection">object selection</param>
		/// <param name="activeObject">object activeObject</param>
		/// <param name="scrollType">optional NetOffice.OWC10Api.Enums.PivotScrollTypeEnum ScrollType = 0</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Select(object selection, object activeObject, object scrollType)
		{
			 Factory.ExecuteMethod(this, "Select", selection, activeObject, scrollType);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="selection">object selection</param>
		/// <param name="activeObject">object activeObject</param>
		/// <param name="scrollType">optional NetOffice.OWC10Api.Enums.PivotScrollTypeEnum ScrollType = 0</param>
		/// <param name="update">optional bool Update = true</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Select(object selection, object activeObject, object scrollType, object update)
		{
			 Factory.ExecuteMethod(this, "Select", selection, activeObject, scrollType, update);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="topic">Int32 topic</param>
		[SupportByVersion("OWC10", 1)]
		public void ShowHelp(Int32 topic)
		{
			 Factory.ExecuteMethod(this, "ShowHelp", topic);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void ShowAbout()
		{
			 Factory.ExecuteMethod(this, "ShowAbout");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="menu">object menu</param>
		[SupportByVersion("OWC10", 1)]
		public void ShowContextMenu(Int32 x, Int32 y, object menu)
		{
			 Factory.ExecuteMethod(this, "ShowContextMenu", x, y, menu);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="initialValue">optional object initialValue</param>
		/// <param name="arrowMode">optional NetOffice.OWC10Api.Enums.PivotArrowModeEnum ArrowMode = 0</param>
		/// <param name="caretPosition">optional NetOffice.OWC10Api.Enums.PivotCaretPositionEnum CaretPosition = 0</param>
		[SupportByVersion("OWC10", 1)]
		public void StartEdit(object initialValue, object arrowMode, object caretPosition)
		{
			 Factory.ExecuteMethod(this, "StartEdit", initialValue, arrowMode, caretPosition);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void StartEdit()
		{
			 Factory.ExecuteMethod(this, "StartEdit");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="initialValue">optional object initialValue</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void StartEdit(object initialValue)
		{
			 Factory.ExecuteMethod(this, "StartEdit", initialValue);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="initialValue">optional object initialValue</param>
		/// <param name="arrowMode">optional NetOffice.OWC10Api.Enums.PivotArrowModeEnum ArrowMode = 0</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void StartEdit(object initialValue, object arrowMode)
		{
			 Factory.ExecuteMethod(this, "StartEdit", initialValue, arrowMode);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="accept">optional bool Accept = true</param>
		[SupportByVersion("OWC10", 1)]
		public void EndEdit(object accept)
		{
			 Factory.ExecuteMethod(this, "EndEdit", accept);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void EndEdit()
		{
			 Factory.ExecuteMethod(this, "EndEdit");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void CancelDragDrop()
		{
			 Factory.ExecuteMethod(this, "CancelDragDrop");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void OkToBindToControlByName()
		{
			 Factory.ExecuteMethod(this, "OkToBindToControlByName");
		}

		#endregion

		#pragma warning restore
	}
}
