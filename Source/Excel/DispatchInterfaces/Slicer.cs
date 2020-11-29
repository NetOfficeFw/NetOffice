using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface Slicer 
	/// SupportByVersion Excel, 14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer"/> </remarks>
	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Slicer : COMObject
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
                    _type = typeof(Slicer);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Slicer(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Slicer(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Slicer(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Slicer(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Slicer(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Slicer(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Slicer() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Slicer(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.Application"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.Creator"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.Parent"/> </remarks>
		[SupportByVersion("Excel", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.Name"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
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
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.Caption"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
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
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.Top"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public Double Top
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Top");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Top", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.Left"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public Double Left
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Left");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Left", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.DisableMoveResizeUI"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public bool DisableMoveResizeUI
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisableMoveResizeUI");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisableMoveResizeUI", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.Width"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public Double Width
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Width");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.Height"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public Double Height
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Height");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Height", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.RowHeight"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public Double RowHeight
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "RowHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RowHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.ColumnWidth"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public Double ColumnWidth
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "ColumnWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ColumnWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.NumberOfColumns"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 NumberOfColumns
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "NumberOfColumns");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NumberOfColumns", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.DisplayHeader"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public bool DisplayHeader
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayHeader");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayHeader", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.Locked"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public bool Locked
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Locked");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Locked", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.SlicerCache"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.SlicerCache SlicerCache
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SlicerCache>(this, "SlicerCache", NetOffice.ExcelApi.SlicerCache.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.SlicerCacheLevel"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.SlicerCacheLevel SlicerCacheLevel
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SlicerCacheLevel>(this, "SlicerCacheLevel", NetOffice.ExcelApi.SlicerCacheLevel.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.Shape"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Shape Shape
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Shape>(this, "Shape", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.Style"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public object Style
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Style");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Style", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.ActiveItem"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.SlicerItem ActiveItem
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SlicerItem>(this, "ActiveItem", NetOffice.ExcelApi.SlicerItem.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.slicer.timelineviewstate"/> </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.TimelineViewState TimelineViewState
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.TimelineViewState>(this, "TimelineViewState", NetOffice.ExcelApi.TimelineViewState.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.slicer.slicercachetype"/> </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Enums.XlSlicerCacheType SlicerCacheType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlSlicerCacheType>(this, "SlicerCacheType");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.Delete"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.Cut"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public void Cut()
		{
			 Factory.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Slicer.Copy"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		#endregion

		#pragma warning restore
	}
}
