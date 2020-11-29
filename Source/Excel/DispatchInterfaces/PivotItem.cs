using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// PivotItem
	/// </summary>
	[SyntaxBypass]
 	public class PivotItem_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PivotItem_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PivotItem_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotItem_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotItem_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotItem_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotItem_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotItem_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotItem_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">optional object index</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.ChildItems"/>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_ChildItems(object index)
		{
			return Factory.ExecuteVariantPropertyGet(this, "ChildItems", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ChildItems
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.ChildItems"/> </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_ChildItems")]
		public object ChildItems(object index)
		{
			return get_ChildItems(index);
		}

		#endregion

		#region Methods

		#endregion
	}

	/// <summary>
	/// DispatchInterface PivotItem 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PivotItem : PivotItem_
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
                    _type = typeof(PivotItem);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PivotItem(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PivotItem(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotItem(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotItem(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotItem(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotItem(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotItem() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotItem(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.Creator"/> </remarks>
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
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.Parent"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotField Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotField>(this, "Parent", NetOffice.ExcelApi.PivotField.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.ChildItems"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object ChildItems
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ChildItems");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.DataRange"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range DataRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "DataRange", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string _Default
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "_Default");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "_Default", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.LabelRange"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range LabelRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "LabelRange", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.Name"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.ParentItem"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotItem ParentItem
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotItem>(this, "ParentItem", NetOffice.ExcelApi.PivotItem.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.ParentShowDetail"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ParentShowDetail
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ParentShowDetail");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.Position"/> </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.ShowDetail"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool ShowDetail
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowDetail");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowDetail", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.SourceName"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object SourceName
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "SourceName");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.Value"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string Value
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Value");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Value", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.Visible"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Visible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Visible");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Visible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.IsCalculated"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool IsCalculated
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsCalculated");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.RecordCount"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 RecordCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RecordCount");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.Formula"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string Formula
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Formula");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Formula", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.Caption"/> </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.DrilledDown"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool DrilledDown
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DrilledDown");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DrilledDown", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.StandardFormula"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public string StandardFormula
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "StandardFormula");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StandardFormula", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.SourceNameStandard"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public string SourceNameStandard
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SourceNameStandard");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.Delete"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotItem.DrillTo"/> </remarks>
		/// <param name="field">string field</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void DrillTo(string field)
		{
			 Factory.ExecuteMethod(this, "DrillTo", field);
		}

		#endregion

		#pragma warning restore
	}
}
