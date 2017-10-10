using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface ContentControl 
	/// SupportByVersion Word, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821215.aspx </remarks>
	[SupportByVersion("Word", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ContentControl : COMObject
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
                    _type = typeof(ContentControl);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ContentControl(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ContentControl(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ContentControl(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ContentControl(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ContentControl(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ContentControl(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ContentControl() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ContentControl(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845327.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193736.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840017.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839788.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Range Range
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Range", NetOffice.WordApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835775.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool LockContentControl
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LockContentControl");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LockContentControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822956.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool LockContents
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LockContents");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LockContents", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193642.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.XMLMapping XMLMapping
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLMapping>(this, "XMLMapping", NetOffice.WordApi.XMLMapping.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192401.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdContentControlType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdContentControlType>(this, "Type");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194660.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.ContentControlListEntries DropdownListEntries
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ContentControlListEntries>(this, "DropdownListEntries", NetOffice.WordApi.ContentControlListEntries.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194687.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.BuildingBlock PlaceholderText
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.BuildingBlock>(this, "PlaceholderText", NetOffice.WordApi.BuildingBlock.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194027.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public string Title
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Title");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Title", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845679.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public string DateDisplayFormat
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DateDisplayFormat");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DateDisplayFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195729.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool MultiLine
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MultiLine");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MultiLine", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822193.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.ContentControl ParentContentControl
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ContentControl>(this, "ParentContentControl", NetOffice.WordApi.ContentControl.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197494.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Temporary
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Temporary");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Temporary", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192748.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public string ID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ID");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191954.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool ShowingPlaceholderText
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowingPlaceholderText");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838048.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdContentControlDateStorageFormat DateStorageFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdContentControlDateStorageFormat>(this, "DateStorageFormat");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DateStorageFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198199.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdBuildingBlockTypes BuildingBlockType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdBuildingBlockTypes>(this, "BuildingBlockType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BuildingBlockType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838162.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public string BuildingBlockCategory
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BuildingBlockCategory");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BuildingBlockCategory", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193973.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdLanguageID DateDisplayLocale
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLanguageID>(this, "DateDisplayLocale");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DateDisplayLocale", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820775.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public object DefaultTextStyle
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultTextStyle");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DefaultTextStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196651.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdCalendarType DateCalendarType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdCalendarType>(this, "DateCalendarType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DateCalendarType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195293.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public string Tag
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Tag");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Tag", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194593.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool Checked
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Checked");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Checked", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227730.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Enums.WdColor Color
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColor>(this, "Color");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Color", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227646.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Enums.WdContentControlAppearance Appearance
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdContentControlAppearance>(this, "Appearance");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Appearance", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229832.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Enums.WdContentControlLevel Level
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdContentControlLevel>(this, "Level");
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231663.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.RepeatingSectionItemColl RepeatingSectionItems
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.RepeatingSectionItemColl>(this, "RepeatingSectionItems", NetOffice.WordApi.RepeatingSectionItemColl.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230582.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public string RepeatingSectionItemTitle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RepeatingSectionItemTitle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RepeatingSectionItemTitle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229741.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public bool AllowInsertDeleteSection
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowInsertDeleteSection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowInsertDeleteSection", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838347.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836258.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void Cut()
		{
			 Factory.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194849.aspx </remarks>
		/// <param name="deleteContents">optional bool DeleteContents = false</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void Delete(object deleteContents)
		{
			 Factory.ExecuteMethod(this, "Delete", deleteContents);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194849.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838691.aspx </remarks>
		/// <param name="buildingBlock">optional NetOffice.WordApi.BuildingBlock BuildingBlock = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Range Range = 0</param>
		/// <param name="text">optional string Text = </param>
		[SupportByVersion("Word", 12,14,15,16)]
        [KnownIssue]
		public virtual void SetPlaceholderText(object buildingBlock, object range, object text)
		{
			 Factory.ExecuteMethod(this, "SetPlaceholderText", buildingBlock, range, text);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838691.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
        [KnownIssue]
        public virtual void SetPlaceholderText()
		{
			 Factory.ExecuteMethod(this, "SetPlaceholderText");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838691.aspx </remarks>
		/// <param name="buildingBlock">optional NetOffice.WordApi.BuildingBlock BuildingBlock = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
        [KnownIssue]
        public virtual void SetPlaceholderText(object buildingBlock)
		{
			 Factory.ExecuteMethod(this, "SetPlaceholderText", buildingBlock);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838691.aspx </remarks>
		/// <param name="buildingBlock">optional NetOffice.WordApi.BuildingBlock BuildingBlock = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Range Range = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
        [KnownIssue]
        public virtual void SetPlaceholderText(object buildingBlock, object range)
		{
			 Factory.ExecuteMethod(this, "SetPlaceholderText", buildingBlock, range);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195895.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void Ungroup()
		{
			 Factory.ExecuteMethod(this, "Ungroup");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197558.aspx </remarks>
		/// <param name="characterNumber">Int32 characterNumber</param>
		/// <param name="font">optional string Font = </param>
		[SupportByVersion("Word", 14,15,16)]
		public void SetCheckedSymbol(Int32 characterNumber, object font)
		{
			 Factory.ExecuteMethod(this, "SetCheckedSymbol", characterNumber, font);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197558.aspx </remarks>
		/// <param name="characterNumber">Int32 characterNumber</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SetCheckedSymbol(Int32 characterNumber)
		{
			 Factory.ExecuteMethod(this, "SetCheckedSymbol", characterNumber);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836974.aspx </remarks>
		/// <param name="characterNumber">Int32 characterNumber</param>
		/// <param name="font">optional string Font = </param>
		[SupportByVersion("Word", 14,15,16)]
		public void SetUncheckedSymbol(Int32 characterNumber, object font)
		{
			 Factory.ExecuteMethod(this, "SetUncheckedSymbol", characterNumber, font);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836974.aspx </remarks>
		/// <param name="characterNumber">Int32 characterNumber</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public void SetUncheckedSymbol(Int32 characterNumber)
		{
			 Factory.ExecuteMethod(this, "SetUncheckedSymbol", characterNumber);
		}

		#endregion

		#pragma warning restore
	}
}
