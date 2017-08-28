using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface ListObject 
	/// SupportByVersion Excel, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197604.aspx </remarks>
	[SupportByVersion("Excel", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ListObject : COMObject
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
                    _type = typeof(ListObject);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ListObject(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ListObject(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ListObject(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ListObject(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ListObject(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ListObject(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ListObject() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ListObject(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822924.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196735.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840235.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public string _Default
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "_Default");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837647.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public bool Active
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Active");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841252.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Range DataBodyRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "DataBodyRange", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839725.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public bool DisplayRightToLeft
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayRightToLeft");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837854.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Range HeaderRowRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "HeaderRowRange", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821198.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Range InsertRowRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "InsertRowRange", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821933.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.ListColumns ListColumns
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ListColumns>(this, "ListColumns", NetOffice.ExcelApi.ListColumns.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834452.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.ListRows ListRows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ListRows>(this, "ListRows", NetOffice.ExcelApi.ListRows.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841184.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
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
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841237.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.QueryTable QueryTable
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.QueryTable>(this, "QueryTable", NetOffice.ExcelApi.QueryTable.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839404.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Range
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Range", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837833.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public bool ShowAutoFilter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowAutoFilter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowAutoFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836501.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public bool ShowTotals
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowTotals");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowTotals", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194428.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlListObjectSourceType SourceType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlListObjectSourceType>(this, "SourceType");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834892.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Range TotalsRowRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "TotalsRowRange", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837420.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public string SharePointURL
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SharePointURL");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835549.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.XmlMap XmlMap
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.XmlMap>(this, "XmlMap", NetOffice.ExcelApi.XmlMap.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193006.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public string DisplayName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DisplayName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836536.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ShowHeaders
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowHeaders");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowHeaders", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836829.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.AutoFilter AutoFilter
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.AutoFilter>(this, "AutoFilter", NetOffice.ExcelApi.AutoFilter.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840453.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public object TableStyle
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TableStyle");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "TableStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194313.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ShowTableStyleFirstColumn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowTableStyleFirstColumn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowTableStyleFirstColumn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821044.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ShowTableStyleLastColumn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowTableStyleLastColumn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowTableStyleLastColumn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840273.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ShowTableStyleRowStripes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowTableStyleRowStripes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowTableStyleRowStripes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196162.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ShowTableStyleColumnStripes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowTableStyleColumnStripes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowTableStyleColumnStripes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836133.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Sort Sort
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sort>(this, "Sort", NetOffice.ExcelApi.Sort.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822164.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public string Comment
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Comment");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Comment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196537.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
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
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198268.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public string Summary
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Summary");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Summary", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230775.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.TableObject TableObject
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.TableObject>(this, "TableObject", NetOffice.ExcelApi.TableObject.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230953.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Slicers Slicers
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Slicers>(this, "Slicers", NetOffice.ExcelApi.Slicers.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231020.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public bool ShowAutoFilterDropDown
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowAutoFilterDropDown");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowAutoFilterDropDown", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839211.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835538.aspx </remarks>
		/// <param name="target">object target</param>
		/// <param name="linkSource">bool linkSource</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public string Publish(object target, bool linkSource)
		{
			return Factory.ExecuteStringMethodGet(this, "Publish", target, linkSource);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834313.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public void Refresh()
		{
			 Factory.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196609.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public void Unlink()
		{
			 Factory.ExecuteMethod(this, "Unlink");
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193017.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public void Unlist()
		{
			 Factory.ExecuteMethod(this, "Unlist");
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="iConflictType">optional NetOffice.ExcelApi.Enums.XlListConflict iConflictType = 0</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public void UpdateChanges(object iConflictType)
		{
			 Factory.ExecuteMethod(this, "UpdateChanges", iConflictType);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public void UpdateChanges()
		{
			 Factory.ExecuteMethod(this, "UpdateChanges");
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838369.aspx </remarks>
		/// <param name="range">NetOffice.ExcelApi.Range range</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public void Resize(NetOffice.ExcelApi.Range range)
		{
			 Factory.ExecuteMethod(this, "Resize", range);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196053.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void ExportToVisio()
		{
			 Factory.ExecuteMethod(this, "ExportToVisio");
		}

		#endregion

		#pragma warning restore
	}
}
