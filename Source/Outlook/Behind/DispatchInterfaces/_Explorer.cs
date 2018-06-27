using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OutlookApi;

namespace NetOffice.OutlookApi.Behind
{
	/// <summary>
	/// DispatchInterface _Explorer 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Explorer : COMObject, NetOffice.OutlookApi._Explorer
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
                    _contractType = typeof(NetOffice.OutlookApi._Explorer);
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
                    _type = typeof(_Explorer);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _Explorer() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869277.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861302.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlObjectClass>(this, "Class");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864406.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NameSpace>(this, "Session");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863296.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(this, "CommandBars", typeof(NetOffice.OfficeApi.CommandBars));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866719.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi.MAPIFolder CurrentFolder
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi.MAPIFolder>(this, "CurrentFolder");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "CurrentFolder", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866079.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual string Caption
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Caption");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861617.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual object CurrentView
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "CurrentView");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "CurrentView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868816.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867221.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual Int32 Left
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Left");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Left", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868726.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.Panes Panes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Panes>(this, "Panes", typeof(NetOffice.OutlookApi.Panes));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861267.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.Selection Selection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Selection>(this, "Selection", typeof(NetOffice.OutlookApi.Selection));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869921.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual Int32 Top
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Top");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Top", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867095.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866771.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlWindowState WindowState
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlWindowState>(this, "WindowState");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "WindowState", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Views
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Views");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869477.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16), ProxyResult]
		public virtual object HTMLDocument
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "HTMLDocument");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868233.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._NavigationPane NavigationPane
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NavigationPane>(this, "NavigationPane");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865605.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._AccountSelector AccountSelector
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._AccountSelector>(this, "AccountSelector");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869315.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._AttachmentSelection AttachmentSelection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._AttachmentSelection>(this, "AttachmentSelection");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231535.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16), ProxyResult]
		public virtual object ActiveInlineResponse
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ActiveInlineResponse");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230742.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16), ProxyResult]
		public virtual object ActiveInlineResponseWordEditor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ActiveInlineResponseWordEditor");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869525.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void Close()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863904.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void Display()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Display");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865063.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void Activate()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869316.aspx </remarks>
		/// <param name="pane">NetOffice.OutlookApi.Enums.OlPane pane</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual bool IsPaneVisible(NetOffice.OutlookApi.Enums.OlPane pane)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsPaneVisible", pane);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863893.aspx </remarks>
		/// <param name="pane">NetOffice.OutlookApi.Enums.OlPane pane</param>
		/// <param name="visible">bool visible</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void ShowPane(NetOffice.OutlookApi.Enums.OlPane pane, bool visible)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowPane", pane, visible);
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="folder">NetOffice.OutlookApi.MAPIFolder folder</param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public virtual void SelectFolder(NetOffice.OutlookApi.MAPIFolder folder)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SelectFolder", folder);
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="folder">NetOffice.OutlookApi.MAPIFolder folder</param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public virtual void DeselectFolder(NetOffice.OutlookApi.MAPIFolder folder)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeselectFolder", folder);
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="folder">NetOffice.OutlookApi.MAPIFolder folder</param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public virtual bool IsFolderSelected(NetOffice.OutlookApi.MAPIFolder folder)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsFolderSelected", folder);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865861.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual void ClearSearch()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ClearSearch");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869309.aspx </remarks>
		/// <param name="query">string query</param>
		/// <param name="searchScope">NetOffice.OutlookApi.Enums.OlSearchScope searchScope</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual void Search(string query, NetOffice.OutlookApi.Enums.OlSearchScope searchScope)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Search", query, searchScope);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868297.aspx </remarks>
		/// <param name="item">object item</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual bool IsItemSelectableInView(object item)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsItemSelectableInView", item);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868732.aspx </remarks>
		/// <param name="item">object item</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void AddToSelection(object item)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddToSelection", item);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869912.aspx </remarks>
		/// <param name="item">object item</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void RemoveFromSelection(object item)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveFromSelection", item);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860458.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void SelectAllItems()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SelectAllItems");
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862680.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void ClearSelection()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ClearSelection");
		}

		#endregion

		#pragma warning restore
	}
}


