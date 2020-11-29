using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _TableView 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _TableView : COMObject
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
                    _type = typeof(_TableView);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _TableView(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _TableView(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _TableView(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _TableView(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _TableView(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _TableView(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _TableView() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _TableView(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.Application"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Application Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.Class"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlObjectClass>(this, "Class");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.Session"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NameSpace>(this, "Session");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.Parent"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.Language"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string Language
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Language");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Language", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.LockUserChanges"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool LockUserChanges
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LockUserChanges");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LockUserChanges", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.Name"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.SaveOption"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlViewSaveOption SaveOption
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlViewSaveOption>(this, "SaveOption");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.Standard"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool Standard
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Standard");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.ViewType"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlViewType ViewType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlViewType>(this, "ViewType");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.XML"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string XML
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "XML");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "XML", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.Filter"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string Filter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Filter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Filter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.ViewFields"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.ViewFields ViewFields
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.ViewFields>(this, "ViewFields", NetOffice.OutlookApi.ViewFields.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.GroupByFields"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.OrderFields GroupByFields
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.OrderFields>(this, "GroupByFields", NetOffice.OutlookApi.OrderFields.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.SortFields"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.OrderFields SortFields
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.OrderFields>(this, "SortFields", NetOffice.OutlookApi.OrderFields.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.MaxLinesInMultiLineView"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public Int32 MaxLinesInMultiLineView
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MaxLinesInMultiLineView");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaxLinesInMultiLineView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.AutomaticGrouping"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool AutomaticGrouping
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutomaticGrouping");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutomaticGrouping", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.DefaultExpandCollapseSetting"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlDefaultExpandCollapseSetting DefaultExpandCollapseSetting
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlDefaultExpandCollapseSetting>(this, "DefaultExpandCollapseSetting");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultExpandCollapseSetting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.AutomaticColumnSizing"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool AutomaticColumnSizing
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutomaticColumnSizing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutomaticColumnSizing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.Multiline"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlMultiLine MultiLine
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlMultiLine>(this, "MultiLine");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MultiLine", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.MultiLineWidth"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public Int32 MultiLineWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MultiLineWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MultiLineWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.AllowInCellEditing"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool AllowInCellEditing
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowInCellEditing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowInCellEditing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.ShowNewItemRow"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool ShowNewItemRow
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowNewItemRow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowNewItemRow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.GridLineStyle"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlGridLineStyle GridLineStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlGridLineStyle>(this, "GridLineStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "GridLineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.ShowItemsInGroups"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool ShowItemsInGroups
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowItemsInGroups");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowItemsInGroups", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.ShowReadingPane"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool ShowReadingPane
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowReadingPane");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowReadingPane", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.HideReadingPaneHeaderInfo"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool HideReadingPaneHeaderInfo
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HideReadingPaneHeaderInfo");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HideReadingPaneHeaderInfo", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool ShowUnreadAndFlaggedMessages
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowUnreadAndFlaggedMessages");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowUnreadAndFlaggedMessages", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.RowFont"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.ViewFont RowFont
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.ViewFont>(this, "RowFont", NetOffice.OutlookApi.ViewFont.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.ColumnFont"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.ViewFont ColumnFont
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.ViewFont>(this, "ColumnFont", NetOffice.OutlookApi.ViewFont.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.AutoPreviewFont"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.ViewFont AutoPreviewFont
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.ViewFont>(this, "AutoPreviewFont", NetOffice.OutlookApi.ViewFont.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.AutoPreview"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlAutoPreview AutoPreview
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlAutoPreview>(this, "AutoPreview");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AutoPreview", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.AutoFormatRules"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.AutoFormatRules AutoFormatRules
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.AutoFormatRules>(this, "AutoFormatRules", NetOffice.OutlookApi.AutoFormatRules.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.ShowConversationByDate"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public bool ShowConversationByDate
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowConversationByDate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowConversationByDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.ShowFullConversations"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public bool ShowFullConversations
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowFullConversations");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowFullConversations", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.AlwaysExpandConversation"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public bool AlwaysExpandConversation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AlwaysExpandConversation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AlwaysExpandConversation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.ShowConversationSendersAboveSubject"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public bool ShowConversationSendersAboveSubject
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowConversationSendersAboveSubject");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowConversationSendersAboveSubject", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.Apply"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Apply()
		{
			 Factory.ExecuteMethod(this, "Apply");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.Copy"/> </remarks>
		/// <param name="name">string name</param>
		/// <param name="saveOption">optional NetOffice.OutlookApi.Enums.OlViewSaveOption saveOption</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.View Copy(string name, object saveOption)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.View>(this, "Copy", NetOffice.OutlookApi.View.LateBindingApiWrapperType, name, saveOption);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.Copy"/> </remarks>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.View Copy(string name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.View>(this, "Copy", NetOffice.OutlookApi.View.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.Delete"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.Reset"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Reset()
		{
			 Factory.ExecuteMethod(this, "Reset");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.Save"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Save()
		{
			 Factory.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.GoToDate"/> </remarks>
		/// <param name="date">DateTime date</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void GoToDate(DateTime date)
		{
			 Factory.ExecuteMethod(this, "GoToDate", date);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.TableView.GetTable"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public NetOffice.OutlookApi.Table GetTable()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Table>(this, "GetTable", NetOffice.OutlookApi.Table.LateBindingApiWrapperType);
		}

		#endregion

		#pragma warning restore
	}
}
