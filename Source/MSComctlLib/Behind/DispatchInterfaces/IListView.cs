using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSComctlLibApi;

namespace NetOffice.MSComctlLibApi.Behind
{
	/// <summary>
	/// DispatchInterface IListView 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IListView : COMObject, NetOffice.MSComctlLibApi.IListView
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
                    _contractType = typeof(NetOffice.MSComctlLibApi.IListView);
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
                    _type = typeof(IListView);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IListView() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.Enums.ListArrangeConstants Arrange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.ListArrangeConstants>(this, "Arrange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Arrange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		public virtual NetOffice.MSComctlLibApi.IColumnHeaders ColumnHeaders
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSComctlLibApi.IColumnHeaders>(this, "ColumnHeaders");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "ColumnHeaders", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSComctlLibApi.IListItem DropHighlight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSComctlLibApi.IListItem>(this, "DropHighlight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "DropHighlight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual bool HideColumnHeaders
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HideColumnHeaders");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HideColumnHeaders", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual bool HideSelection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HideSelection");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HideSelection", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSComctlLib", 6), ProxyResult]
		public virtual object Icons
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Icons");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Icons", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSComctlLibApi.IListItems ListItems
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSComctlLibApi.IListItems>(this, "ListItems");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "ListItems", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.Enums.ListLabelEditConstants LabelEdit
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.ListLabelEditConstants>(this, "LabelEdit");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "LabelEdit", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual bool LabelWrap
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "LabelWrap");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LabelWrap", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6), NativeResult]
		public virtual stdole.Picture MouseIcon
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MouseIcon", paramsArray);
                return returnItem as stdole.Picture;
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
		public virtual NetOffice.MSComctlLibApi.Enums.MousePointerConstants MousePointer
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.MousePointerConstants>(this, "MousePointer");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "MousePointer", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual bool MultiSelect
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MultiSelect");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MultiSelect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.MSComctlLibApi.IListItem SelectedItem
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSComctlLibApi.IListItem>(this, "SelectedItem");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "SelectedItem", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSComctlLib", 6), ProxyResult]
		public virtual object SmallIcons
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "SmallIcons");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "SmallIcons", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual bool Sorted
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Sorted");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Sorted", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual Int16 SortKey
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "SortKey");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SortKey", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.Enums.ListSortOrderConstants SortOrder
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.ListSortOrderConstants>(this, "SortOrder");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SortOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.Enums.ListViewConstants View
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.ListViewConstants>(this, "View");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "View", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.Enums.OLEDragConstants OLEDragMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.OLEDragConstants>(this, "OLEDragMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "OLEDragMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.Enums.OLEDropConstants OLEDropMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.OLEDropConstants>(this, "OLEDropMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "OLEDropMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.Enums.AppearanceConstants Appearance
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.AppearanceConstants>(this, "Appearance");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Appearance", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual Int32 BackColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BackColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.Enums.BorderStyleConstants BorderStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.BorderStyleConstants>(this, "BorderStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "BorderStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual bool Enabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Enabled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Enabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6), NativeResult]
		public virtual stdole.Font Font
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Font", paramsArray);
                return returnItem as stdole.Font;
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
		public virtual Int32 ForeColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ForeColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 hWnd
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "hWnd");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "hWnd", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual bool AllowColumnReorder
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowColumnReorder");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowColumnReorder", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual bool Checkboxes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Checkboxes");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Checkboxes", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual bool FlatScrollBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FlatScrollBar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FlatScrollBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual bool FullRowSelect
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FullRowSelect");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FullRowSelect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual bool GridLines
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "GridLines");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual bool HotTracking
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HotTracking");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HotTracking", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual bool HoverSelection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HoverSelection");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HoverSelection", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6), NativeResult]
		public virtual stdole.Picture Picture
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Picture", paramsArray);
                return returnItem as stdole.Picture;
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
		public virtual NetOffice.MSComctlLibApi.Enums.ListPictureAlignmentConstants PictureAlignment
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.ListPictureAlignmentConstants>(this, "PictureAlignment");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PictureAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSComctlLib", 6), ProxyResult]
		public virtual object ColumnHeaderIcons
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ColumnHeaderIcons");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "ColumnHeaderIcons", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.Enums.ListTextBackgroundConstants TextBackground
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.ListTextBackgroundConstants>(this, "TextBackground");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "TextBackground", value);
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
		[BaseResult]
		public virtual NetOffice.MSComctlLibApi.IListItem FindItem(string sz, object where, object index, object fPartial)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IListItem>(this, "FindItem", sz, where, index, fPartial);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="sz">string sz</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.IListItem FindItem(string sz)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IListItem>(this, "FindItem", sz);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="sz">string sz</param>
		/// <param name="where">optional object where</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.IListItem FindItem(string sz, object where)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IListItem>(this, "FindItem", sz, where);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="sz">string sz</param>
		/// <param name="where">optional object where</param>
		/// <param name="index">optional object index</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.IListItem FindItem(string sz, object where, object index)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IListItem>(this, "FindItem", sz, where, index);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		[BaseResult]
		public virtual NetOffice.MSComctlLibApi.IListItem GetFirstVisible()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IListItem>(this, "GetFirstVisible");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="x">Single x</param>
		/// <param name="y">Single y</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.IListItem HitTest(Single x, Single y)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSComctlLibApi.IListItem>(this, "HitTest", x, y);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void StartLabelEdit()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "StartLabelEdit");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void OLEDrag()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OLEDrag");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void Refresh()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void AboutBox()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AboutBox");
		}

		#endregion

		#pragma warning restore
	}
}

