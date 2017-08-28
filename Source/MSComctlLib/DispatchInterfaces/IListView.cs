using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi
{
	/// <summary>
	/// DispatchInterface IListView 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IListView : COMObject
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
                    _type = typeof(IListView);
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IListView(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListView(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListView(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListView(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListView(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListView() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IListView(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.Enums.ListArrangeConstants Arrange
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.ListArrangeConstants>(this, "Arrange");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Arrange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IColumnHeaders ColumnHeaders
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSComctlLibApi.IColumnHeaders>(this, "ColumnHeaders", NetOffice.MSComctlLibApi.IColumnHeaders.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "ColumnHeaders", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSComctlLibApi.IListItem DropHighlight
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSComctlLibApi.IListItem>(this, "DropHighlight", NetOffice.MSComctlLibApi.IListItem.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "DropHighlight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public bool HideColumnHeaders
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HideColumnHeaders");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HideColumnHeaders", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public bool HideSelection
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HideSelection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HideSelection", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSComctlLib", 6), ProxyResult]
		public object Icons
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Icons");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Icons", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSComctlLibApi.IListItems ListItems
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSComctlLibApi.IListItems>(this, "ListItems", NetOffice.MSComctlLibApi.IListItems.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "ListItems", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.Enums.ListLabelEditConstants LabelEdit
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.ListLabelEditConstants>(this, "LabelEdit");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LabelEdit", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public bool LabelWrap
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LabelWrap");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LabelWrap", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public stdole.Picture MouseIcon
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MouseIcon", paramsArray);
				stdole.Picture newObject = Factory.CreateObjectFromComProxy(this,returnItem) as stdole.Picture;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MouseIcon", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.Enums.MousePointerConstants MousePointer
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.MousePointerConstants>(this, "MousePointer");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MousePointer", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public bool MultiSelect
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MultiSelect");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MultiSelect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSComctlLibApi.IListItem SelectedItem
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSComctlLibApi.IListItem>(this, "SelectedItem", NetOffice.MSComctlLibApi.IListItem.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "SelectedItem", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSComctlLib", 6), ProxyResult]
		public object SmallIcons
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "SmallIcons");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "SmallIcons", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public bool Sorted
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Sorted");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Sorted", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public Int16 SortKey
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "SortKey");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SortKey", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.Enums.ListSortOrderConstants SortOrder
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.ListSortOrderConstants>(this, "SortOrder");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SortOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.Enums.ListViewConstants View
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.ListViewConstants>(this, "View");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "View", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.Enums.OLEDragConstants OLEDragMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.OLEDragConstants>(this, "OLEDragMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "OLEDragMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.Enums.OLEDropConstants OLEDropMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.OLEDropConstants>(this, "OLEDropMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "OLEDropMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.Enums.AppearanceConstants Appearance
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.AppearanceConstants>(this, "Appearance");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Appearance", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public Int32 BackColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BackColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.Enums.BorderStyleConstants BorderStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.BorderStyleConstants>(this, "BorderStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BorderStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public bool Enabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Enabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Enabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public stdole.Font Font
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Font", paramsArray);
				stdole.Font newObject = Factory.CreateObjectFromComProxy(this,returnItem) as stdole.Font;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Font", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public Int32 ForeColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ForeColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 hWnd
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "hWnd");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "hWnd", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public bool AllowColumnReorder
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowColumnReorder");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowColumnReorder", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public bool Checkboxes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Checkboxes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Checkboxes", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public bool FlatScrollBar
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FlatScrollBar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FlatScrollBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public bool FullRowSelect
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FullRowSelect");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FullRowSelect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public bool GridLines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "GridLines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public bool HotTracking
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HotTracking");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HotTracking", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public bool HoverSelection
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HoverSelection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HoverSelection", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public stdole.Picture Picture
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Picture", paramsArray);
				stdole.Picture newObject = Factory.CreateObjectFromComProxy(this,returnItem) as stdole.Picture;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Picture", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.Enums.ListPictureAlignmentConstants PictureAlignment
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.ListPictureAlignmentConstants>(this, "PictureAlignment");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PictureAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSComctlLib", 6), ProxyResult]
		public object ColumnHeaderIcons
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ColumnHeaderIcons");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "ColumnHeaderIcons", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.Enums.ListTextBackgroundConstants TextBackground
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.ListTextBackgroundConstants>(this, "TextBackground");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TextBackground", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="sz">string sz</param>
		/// <param name="where">optional object where</param>
		/// <param name="index">optional object index</param>
		/// <param name="fPartial">optional object fPartial</param>
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IListItem FindItem(string sz, object where, object index, object fPartial)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSComctlLibApi.IListItem>(this, "FindItem", NetOffice.MSComctlLibApi.IListItem.LateBindingApiWrapperType, sz, where, index, fPartial);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="sz">string sz</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IListItem FindItem(string sz)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSComctlLibApi.IListItem>(this, "FindItem", NetOffice.MSComctlLibApi.IListItem.LateBindingApiWrapperType, sz);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="sz">string sz</param>
		/// <param name="where">optional object where</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IListItem FindItem(string sz, object where)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSComctlLibApi.IListItem>(this, "FindItem", NetOffice.MSComctlLibApi.IListItem.LateBindingApiWrapperType, sz, where);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="sz">string sz</param>
		/// <param name="where">optional object where</param>
		/// <param name="index">optional object index</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IListItem FindItem(string sz, object where, object index)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSComctlLibApi.IListItem>(this, "FindItem", NetOffice.MSComctlLibApi.IListItem.LateBindingApiWrapperType, sz, where, index);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IListItem GetFirstVisible()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSComctlLibApi.IListItem>(this, "GetFirstVisible", NetOffice.MSComctlLibApi.IListItem.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="x">Single x</param>
		/// <param name="y">Single y</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IListItem HitTest(Single x, Single y)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSComctlLibApi.IListItem>(this, "HitTest", NetOffice.MSComctlLibApi.IListItem.LateBindingApiWrapperType, x, y);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public void StartLabelEdit()
		{
			 Factory.ExecuteMethod(this, "StartLabelEdit");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public void OLEDrag()
		{
			 Factory.ExecuteMethod(this, "OLEDrag");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public void Refresh()
		{
			 Factory.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSComctlLib", 6)]
		public void AboutBox()
		{
			 Factory.ExecuteMethod(this, "AboutBox");
		}

		#endregion

		#pragma warning restore
	}
}



