using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface Shape 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Shape : COMObject
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
                    _type = typeof(Shape);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Shape(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Shape(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Shape(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Shape(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Shape(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Shape(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Shape() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Shape(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", NetOffice.PublisherApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Adjustments Adjustments
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Adjustments>(this, "Adjustments", NetOffice.PublisherApi.Adjustments.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoAutoShapeType AutoShapeType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoAutoShapeType>(this, "AutoShapeType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AutoShapeType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.CalloutFormat Callout
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.CalloutFormat>(this, "Callout", NetOffice.PublisherApi.CalloutFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 ConnectionSiteCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ConnectionSiteCount");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Connector
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Connector");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ConnectorFormat ConnectorFormat
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ConnectorFormat>(this, "ConnectorFormat", NetOffice.PublisherApi.ConnectorFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.FillFormat Fill
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.FillFormat>(this, "Fill", NetOffice.PublisherApi.FillFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.GroupShapes GroupItems
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.GroupShapes>(this, "GroupItems", NetOffice.PublisherApi.GroupShapes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object Height
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Height");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Height", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState HorizontalFlip
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "HorizontalFlip");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object Left
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Left");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Left", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.LineFormat Line
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.LineFormat>(this, "Line", NetOffice.PublisherApi.LineFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState LockAspectRatio
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "LockAspectRatio");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LockAspectRatio", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeNodes Nodes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ShapeNodes>(this, "Nodes", NetOffice.PublisherApi.ShapeNodes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single Rotation
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Rotation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Rotation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.PictureFormat PictureFormat
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.PictureFormat>(this, "PictureFormat", NetOffice.PublisherApi.PictureFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShadowFormat Shadow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ShadowFormat>(this, "Shadow", NetOffice.PublisherApi.ShadowFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextEffectFormat TextEffect
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.TextEffectFormat>(this, "TextEffect", NetOffice.PublisherApi.TextEffectFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextFrame TextFrame
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.TextFrame>(this, "TextFrame", NetOffice.PublisherApi.TextFrame.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ThreeDFormat ThreeD
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ThreeDFormat>(this, "ThreeD", NetOffice.PublisherApi.ThreeDFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object Top
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Top");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Top", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbShapeType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbShapeType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState VerticalFlip
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "VerticalFlip");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object Vertices
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Vertices");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public object Width
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Width");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Width", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 ZOrderPosition
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ZOrderPosition");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string AlternativeText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AlternativeText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AlternativeText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState HasTable
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "HasTable");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState HasTextFrame
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "HasTextFrame");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Hyperlink Hyperlink
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Hyperlink>(this, "Hyperlink", NetOffice.PublisherApi.Hyperlink.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbWizardTag WizardTag
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbWizardTag>(this, "WizardTag");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "WizardTag", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 ID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ID");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.LinkFormat LinkFormat
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.LinkFormat>(this, "LinkFormat", NetOffice.PublisherApi.LinkFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.OLEFormat OLEFormat
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.OLEFormat>(this, "OLEFormat", NetOffice.PublisherApi.OLEFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Table Table
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Table>(this, "Table", NetOffice.PublisherApi.Table.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Tags Tags
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Tags>(this, "Tags", NetOffice.PublisherApi.Tags.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoBlackWhiteMode BlackWhiteMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoBlackWhiteMode>(this, "BlackWhiteMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BlackWhiteMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool IsGroupMember
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsGroupMember");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape ParentGroupShape
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Shape>(this, "ParentGroupShape", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 WizardTagInstance
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "WizardTagInstance");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WizardTagInstance", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.WebCommandButton WebCommandButton
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.WebCommandButton>(this, "WebCommandButton", NetOffice.PublisherApi.WebCommandButton.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.WebListBox WebListBox
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.WebListBox>(this, "WebListBox", NetOffice.PublisherApi.WebListBox.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.WebTextBox WebTextBox
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.WebTextBox>(this, "WebTextBox", NetOffice.PublisherApi.WebTextBox.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.WebOptionButton WebOptionButton
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.WebOptionButton>(this, "WebOptionButton", NetOffice.PublisherApi.WebOptionButton.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.WebCheckBox WebCheckBox
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.WebCheckBox>(this, "WebCheckBox", NetOffice.PublisherApi.WebCheckBox.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Wizard Wizard
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Wizard>(this, "Wizard", NetOffice.PublisherApi.Wizard.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.WrapFormat TextWrap
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.WrapFormat>(this, "TextWrap", NetOffice.PublisherApi.WrapFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.WebComponentFormat WebComponentFormat
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.WebComponentFormat>(this, "WebComponentFormat", NetOffice.OfficeApi.WebComponentFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.BorderArtFormat BorderArt
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.BorderArtFormat>(this, "BorderArt", NetOffice.PublisherApi.BorderArtFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string WebNavigationBarSetName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "WebNavigationBarSetName");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.CatalogMergeShapes CatalogMergeItems
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.CatalogMergeShapes>(this, "CatalogMergeItems", NetOffice.PublisherApi.CatalogMergeShapes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState IsInline
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "IsInline");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.TextRange InlineTextRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.TextRange>(this, "InlineTextRange", NetOffice.PublisherApi.TextRange.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbInlineAlignment InlineAlignment
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbInlineAlignment>(this, "InlineAlignment");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "InlineAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState IsExcess
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "IsExcess");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		public NetOffice.PublisherApi.SoftEdgeFormat SoftEdge
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.SoftEdgeFormat>(this, "SoftEdge", NetOffice.PublisherApi.SoftEdgeFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		public NetOffice.PublisherApi.GlowFormat Glow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.GlowFormat>(this, "Glow", NetOffice.PublisherApi.GlowFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		public NetOffice.PublisherApi.ReflectionFormat Reflection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ReflectionFormat>(this, "Reflection", NetOffice.PublisherApi.ReflectionFormat.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Apply()
		{
			 Factory.ExecuteMethod(this, "Apply");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Cut()
		{
			 Factory.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape Duplicate()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "Duplicate", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="flipCmd">NetOffice.OfficeApi.Enums.MsoFlipCmd flipCmd</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Flip(NetOffice.OfficeApi.Enums.MsoFlipCmd flipCmd)
		{
			 Factory.ExecuteMethod(this, "Flip", flipCmd);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbUnitType unit</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single GetLeft(NetOffice.PublisherApi.Enums.PbUnitType unit)
		{
			return Factory.ExecuteSingleMethodGet(this, "GetLeft", unit);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbUnitType unit</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single GetTop(NetOffice.PublisherApi.Enums.PbUnitType unit)
		{
			return Factory.ExecuteSingleMethodGet(this, "GetTop", unit);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbUnitType unit</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single GetHeight(NetOffice.PublisherApi.Enums.PbUnitType unit)
		{
			return Factory.ExecuteSingleMethodGet(this, "GetHeight", unit);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbUnitType unit</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public Single GetWidth(NetOffice.PublisherApi.Enums.PbUnitType unit)
		{
			return Factory.ExecuteSingleMethodGet(this, "GetWidth", unit);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="increment">object increment</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void IncrementLeft(object increment)
		{
			 Factory.ExecuteMethod(this, "IncrementLeft", increment);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void IncrementRotation(Single increment)
		{
			 Factory.ExecuteMethod(this, "IncrementRotation", increment);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="increment">object increment</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void IncrementTop(object increment)
		{
			 Factory.ExecuteMethod(this, "IncrementTop", increment);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void PickUp()
		{
			 Factory.ExecuteMethod(this, "PickUp");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void RerouteConnections()
		{
			 Factory.ExecuteMethod(this, "RerouteConnections");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="factor">Single factor</param>
		/// <param name="relativeToOriginalSize">NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize</param>
		/// <param name="fScale">optional NetOffice.OfficeApi.Enums.MsoScaleFrom fScale</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void ScaleHeight(Single factor, NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize, object fScale)
		{
			 Factory.ExecuteMethod(this, "ScaleHeight", factor, relativeToOriginalSize, fScale);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="factor">Single factor</param>
		/// <param name="relativeToOriginalSize">NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ScaleHeight(Single factor, NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize)
		{
			 Factory.ExecuteMethod(this, "ScaleHeight", factor, relativeToOriginalSize);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="factor">Single factor</param>
		/// <param name="relativeToOriginalSize">NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize</param>
		/// <param name="fScale">optional NetOffice.OfficeApi.Enums.MsoScaleFrom fScale</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void ScaleWidth(Single factor, NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize, object fScale)
		{
			 Factory.ExecuteMethod(this, "ScaleWidth", factor, relativeToOriginalSize, fScale);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="factor">Single factor</param>
		/// <param name="relativeToOriginalSize">NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ScaleWidth(Single factor, NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize)
		{
			 Factory.ExecuteMethod(this, "ScaleWidth", factor, relativeToOriginalSize);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="replace">optional object replace</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Select(object replace)
		{
			 Factory.ExecuteMethod(this, "Select", replace);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void SetShapesDefaultProperties()
		{
			 Factory.ExecuteMethod(this, "SetShapesDefaultProperties");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange Ungroup()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ShapeRange>(this, "Ungroup", NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="zOrderCmd">NetOffice.OfficeApi.Enums.MsoZOrderCmd zOrderCmd</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void ZOrder(NetOffice.OfficeApi.Enums.MsoZOrderCmd zOrderCmd)
		{
			 Factory.ExecuteMethod(this, "ZOrder", zOrderCmd);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void AddToCatalogMergeArea()
		{
			 Factory.ExecuteMethod(this, "AddToCatalogMergeArea");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void RemoveFromCatalogMergeArea()
		{
			 Factory.ExecuteMethod(this, "RemoveFromCatalogMergeArea");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void RemoveCatalogMergeArea()
		{
			 Factory.ExecuteMethod(this, "RemoveCatalogMergeArea");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange range</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void MoveIntoTextFlow(NetOffice.PublisherApi.TextRange range)
		{
			 Factory.ExecuteMethod(this, "MoveIntoTextFlow", range);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void MoveOutOfTextFlow()
		{
			 Factory.ExecuteMethod(this, "MoveOutOfTextFlow");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="pbResolution">optional NetOffice.PublisherApi.Enums.PbPictureResolution pbResolution = 0</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void SaveAsPicture(string filename, object pbResolution)
		{
			 Factory.ExecuteMethod(this, "SaveAsPicture", filename, pbResolution);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void SaveAsPicture(string filename)
		{
			 Factory.ExecuteMethod(this, "SaveAsPicture", filename);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="page">Int32 page</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void MoveToPage(Int32 page, object left, object top)
		{
			 Factory.ExecuteMethod(this, "MoveToPage", page, left, top);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="page">Int32 page</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void MoveToPage(Int32 page)
		{
			 Factory.ExecuteMethod(this, "MoveToPage", page);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="page">Int32 page</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void MoveToPage(Int32 page, object left)
		{
			 Factory.ExecuteMethod(this, "MoveToPage", page, left);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="style">NetOffice.PublisherApi.CaptionStyle style</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape SetCaption(NetOffice.PublisherApi.CaptionStyle style)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "SetCaption", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, style);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.BuildingBlock SaveAsBuildingBlock(string name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.BuildingBlock>(this, "SaveAsBuildingBlock", NetOffice.PublisherApi.BuildingBlock.LateBindingApiWrapperType, name);
		}

		#endregion

		#pragma warning restore
	}
}
