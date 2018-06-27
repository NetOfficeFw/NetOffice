using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface IPivotControl 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IPivotControl : COMObject, NetOffice.OWC10Api.IPivotControl
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
                    _contractType = typeof(NetOffice.OWC10Api.IPivotControl);
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
                    _type = typeof(IPivotControl);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IPivotControl() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotView ActiveView
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotView>(this, "ActiveView", typeof(NetOffice.OWC10Api.PivotView));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		public virtual object Selection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Selection");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Selection", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string DataMember
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DataMember");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DataMember", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotData ActiveData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotData>(this, "ActiveData", typeof(NetOffice.OWC10Api.PivotData));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string Version
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool HasDetails
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasDetails");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool DisplayToolbar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayToolbar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayToolbar", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AllowGrouping
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowGrouping");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowGrouping", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AllowFiltering
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowFiltering");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowFiltering", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AllowDetails
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowDetails");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowDetails", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AllowPropertyToolbox
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowPropertyToolbox");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowPropertyToolbox", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AllowCustomOrdering
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowCustomOrdering");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowCustomOrdering", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AutoFit
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoFit");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoFit", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.MSDATASRCApi.DataSource DataSource
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSDATASRCApi.DataSource>(this, "DataSource", typeof(NetOffice.MSDATASRCApi.DataSource));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "DataSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual object BackColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BackColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool DisplayExpandIndicator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayExpandIndicator");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayExpandIndicator", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool RightToLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "RightToLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RightToLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 MaxWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "MaxWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MaxWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 MaxHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "MaxHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MaxHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Width
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Width");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Height
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Height");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Height", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string XMLData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "XMLData");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "XMLData", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool DisplayPropertyToolbox
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayPropertyToolbox");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayPropertyToolbox", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool DisplayFieldList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayFieldList");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayFieldList", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Constants
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Constants");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 MajorVersion
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "MajorVersion");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string MinorVersion
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "MinorVersion");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string BuildNumber
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BuildNumber");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string ConnectionString
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ConnectionString");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ConnectionString", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string CommandText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CommandText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CommandText", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.ProviderType ProviderType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ProviderType>(this, "ProviderType");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.OWC10Api.Enums.PivotTableMemberExpandEnum MemberExpand
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotTableMemberExpandEnum>(this, "MemberExpand");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "MemberExpand", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.ADODBApi.Connection Connection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Connection>(this, "Connection", typeof(NetOffice.ADODBApi.Connection));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Connection", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string RevisionNumber
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RevisionNumber");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool DisplayAlerts
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayAlerts");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayAlerts", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object DataMemberStrings
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DataMemberStrings");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.OWC10Api.PivotClassFactory ClassFactory
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotClassFactory>(this, "ClassFactory", typeof(NetOffice.OWC10Api.PivotClassFactory));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "ClassFactory", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Left
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Left");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Top
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Top");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Hwnd
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Hwnd");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		public virtual object ActiveObject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ActiveObject");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "ActiveObject", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.OCCommands Commands
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.OCCommands>(this, "Commands", typeof(NetOffice.OWC10Api.OCCommands));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool UserMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UserMode");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string DataMemberCaption
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DataMemberCaption");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DataMemberCaption", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object DataSourceEx
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "DataSourceEx");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "DataSourceEx", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool IsDirty
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsDirty");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IsDirty", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string CubeProvider
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CubeProvider");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CubeProvider", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string SelectionType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SelectionType");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool DisplayScreenTips
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayScreenTips");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayScreenTips", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool ViewOnlyMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ViewOnlyMode");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool DisplayDesignTimeUI
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayDesignTimeUI");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayDesignTimeUI", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public virtual NetOffice.MSComctlLibApi.IToolbar Toolbar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSComctlLibApi.IToolbar>(this, "Toolbar");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.PivotEditModeEnum EditMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotEditModeEnum>(this, "EditMode");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string HTMLData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HTMLData");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string DataSourceName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DataSourceName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DataSourceName", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool DisplayBranding
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayBranding");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayBranding", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool DisplayOfficeLogo
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayOfficeLogo");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayOfficeLogo", value);
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
		public virtual void Export(object filename, object action)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Export", filename, action);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void Export()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Export");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = </param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void Export(object filename)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Export", filename);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void Refresh()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = pivot.gif</param>
		/// <param name="filterName">optional string FilterName = gif</param>
		/// <param name="width">optional Int32 Width = 1024</param>
		/// <param name="height">optional Int32 Height = 1024</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void ExportPicture(object filename, object filterName, object width, object height)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportPicture", filename, filterName, width, height);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void ExportPicture()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportPicture");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = pivot.gif</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void ExportPicture(object filename)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportPicture", filename);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = pivot.gif</param>
		/// <param name="filterName">optional string FilterName = gif</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void ExportPicture(object filename, object filterName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportPicture", filename, filterName);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="filename">optional string Filename = pivot.gif</param>
		/// <param name="filterName">optional string FilterName = gif</param>
		/// <param name="width">optional Int32 Width = 1024</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void ExportPicture(object filename, object filterName, object width)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportPicture", filename, filterName, width);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public virtual void LocateDataSource()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LocateDataSource");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="selection">optional object Selection = null (Nothing in visual basic)</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void Copy(object selection)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Copy", selection);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void Copy()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Copy");
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
		public virtual void DoDragDrop(NetOffice.OWC10Api.DropSource source, object dragItem, NetOffice.OWC10Api.DropTarget target, Int32 dwLegalEffect)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DoDragDrop", source, dragItem, target, dwLegalEffect);
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
		public virtual void Select(object selection, object activeObject, object scrollType, object update, object notify)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Select", new object[]{ selection, activeObject, scrollType, update, notify });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="selection">object selection</param>
		/// <param name="activeObject">object activeObject</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void Select(object selection, object activeObject)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Select", selection, activeObject);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="selection">object selection</param>
		/// <param name="activeObject">object activeObject</param>
		/// <param name="scrollType">optional NetOffice.OWC10Api.Enums.PivotScrollTypeEnum ScrollType = 0</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void Select(object selection, object activeObject, object scrollType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Select", selection, activeObject, scrollType);
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
		public virtual void Select(object selection, object activeObject, object scrollType, object update)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Select", selection, activeObject, scrollType, update);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="topic">Int32 topic</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void ShowHelp(Int32 topic)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowHelp", topic);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void ShowAbout()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowAbout");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="menu">object menu</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void ShowContextMenu(Int32 x, Int32 y, object menu)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowContextMenu", x, y, menu);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="initialValue">optional object initialValue</param>
		/// <param name="arrowMode">optional NetOffice.OWC10Api.Enums.PivotArrowModeEnum ArrowMode = 0</param>
		/// <param name="caretPosition">optional NetOffice.OWC10Api.Enums.PivotCaretPositionEnum CaretPosition = 0</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void StartEdit(object initialValue, object arrowMode, object caretPosition)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "StartEdit", initialValue, arrowMode, caretPosition);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void StartEdit()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "StartEdit");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="initialValue">optional object initialValue</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void StartEdit(object initialValue)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "StartEdit", initialValue);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="initialValue">optional object initialValue</param>
		/// <param name="arrowMode">optional NetOffice.OWC10Api.Enums.PivotArrowModeEnum ArrowMode = 0</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void StartEdit(object initialValue, object arrowMode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "StartEdit", initialValue, arrowMode);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="accept">optional bool Accept = true</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void EndEdit(object accept)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EndEdit", accept);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void EndEdit()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EndEdit");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public virtual void CancelDragDrop()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CancelDragDrop");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void OkToBindToControlByName()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OkToBindToControlByName");
		}

		#endregion

		#pragma warning restore
	}
}


