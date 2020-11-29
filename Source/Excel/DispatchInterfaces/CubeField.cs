using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface CubeField 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class CubeField : COMObject
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
                    _type = typeof(CubeField);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public CubeField(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CubeField(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CubeField(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CubeField(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CubeField(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CubeField(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CubeField() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CubeField(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.Application"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.Creator"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.Parent"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.CubeFieldType"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCubeFieldType CubeFieldType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCubeFieldType>(this, "CubeFieldType");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.Caption"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string Caption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Caption");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Caption", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.Name"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.Value"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string Value
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Value");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.Orientation"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlPivotFieldOrientation Orientation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPivotFieldOrientation>(this, "Orientation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Orientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.Position"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Position
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Position");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Position", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.TreeviewControl"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.TreeviewControl TreeviewControl
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.TreeviewControl>(this, "TreeviewControl", NetOffice.ExcelApi.TreeviewControl.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.DragToColumn"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool DragToColumn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DragToColumn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DragToColumn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.DragToHide"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool DragToHide
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DragToHide");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DragToHide", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.DragToPage"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool DragToPage
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DragToPage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DragToPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.DragToRow"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool DragToRow
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DragToRow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DragToRow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.DragToData"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool DragToData
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DragToData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DragToData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 HiddenLevels
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HiddenLevels");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HiddenLevels", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.HasMemberProperties"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool HasMemberProperties
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasMemberProperties");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.LayoutForm"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlLayoutFormType LayoutForm
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlLayoutFormType>(this, "LayoutForm");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LayoutForm", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.PivotFields"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotFields PivotFields
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotFields>(this, "PivotFields", NetOffice.ExcelApi.PivotFields.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.EnableMultiplePageItems"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool EnableMultiplePageItems
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableMultiplePageItems");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableMultiplePageItems", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.LayoutSubtotalLocation"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.xLSubtototalLocationType LayoutSubtotalLocation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.xLSubtototalLocationType>(this, "LayoutSubtotalLocation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LayoutSubtotalLocation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.ShowInFieldList"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool ShowInFieldList
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowInFieldList");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowInFieldList", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string _Caption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "_Caption");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.IncludeNewItemsInFilter"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool IncludeNewItemsInFilter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IncludeNewItemsInFilter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IncludeNewItemsInFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.CubeFieldSubType"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCubeFieldSubType CubeFieldSubType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCubeFieldSubType>(this, "CubeFieldSubType");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.AllItemsVisible"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool AllItemsVisible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllItemsVisible");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.CurrentPageName"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public string CurrentPageName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CurrentPageName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CurrentPageName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.IsDate"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool IsDate
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsDate");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.FlattenHierarchies"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public bool FlattenHierarchies
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FlattenHierarchies");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FlattenHierarchies", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.HierarchizeDistinct"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public bool HierarchizeDistinct
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HierarchizeDistinct");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HierarchizeDistinct", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.AddMemberPropertyField"/> </remarks>
		/// <param name="property">string property</param>
		/// <param name="propertyOrder">optional object propertyOrder</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void AddMemberPropertyField(string property, object propertyOrder)
		{
			 Factory.ExecuteMethod(this, "AddMemberPropertyField", property, propertyOrder);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.AddMemberPropertyField"/> </remarks>
		/// <param name="property">string property</param>
		/// <param name="propertyOrder">optional object propertyOrder</param>
		/// <param name="propertyDisplayedIn">optional object propertyDisplayedIn</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void AddMemberPropertyField(string property, object propertyOrder, object propertyDisplayedIn)
		{
			 Factory.ExecuteMethod(this, "AddMemberPropertyField", property, propertyOrder, propertyDisplayedIn);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.AddMemberPropertyField"/> </remarks>
		/// <param name="property">string property</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void AddMemberPropertyField(string property)
		{
			 Factory.ExecuteMethod(this, "AddMemberPropertyField", property);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.Delete"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="property">string property</param>
		/// <param name="propertyOrder">optional object propertyOrder</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void _AddMemberPropertyField(string property, object propertyOrder)
		{
			 Factory.ExecuteMethod(this, "_AddMemberPropertyField", property, propertyOrder);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="property">string property</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void _AddMemberPropertyField(string property)
		{
			 Factory.ExecuteMethod(this, "_AddMemberPropertyField", property);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.ClearManualFilter"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ClearManualFilter()
		{
			 Factory.ExecuteMethod(this, "ClearManualFilter");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.CubeField.CreatePivotFields"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void CreatePivotFields()
		{
			 Factory.ExecuteMethod(this, "CreatePivotFields");
		}

		#endregion

		#pragma warning restore
	}
}
