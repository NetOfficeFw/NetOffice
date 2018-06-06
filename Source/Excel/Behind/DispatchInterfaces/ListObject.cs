using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
	/// <summary>
	/// DispatchInterface ListObject 
	/// SupportByVersion Excel, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197604.aspx </remarks>
	[SupportByVersion("Excel", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ListObject : COMObject, NetOffice.ExcelApi.ListObject
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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ListObject() : base()
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
		public virtual NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196735.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
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
		public virtual object Parent
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
		public virtual string _Default
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
		public virtual bool Active
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
		public virtual NetOffice.ExcelApi.Range DataBodyRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "DataBodyRange", typeof(NetOffice.ExcelApi.Range));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839725.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual bool DisplayRightToLeft
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
		public virtual NetOffice.ExcelApi.Range HeaderRowRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "HeaderRowRange", typeof(NetOffice.ExcelApi.Range));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821198.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Range InsertRowRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "InsertRowRange", typeof(NetOffice.ExcelApi.Range));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821933.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.ListColumns ListColumns
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ListColumns>(this, "ListColumns", typeof(NetOffice.ExcelApi.ListColumns));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834452.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.ListRows ListRows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ListRows>(this, "ListRows", typeof(NetOffice.ExcelApi.ListRows));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841184.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual string Name
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
		public virtual NetOffice.ExcelApi.QueryTable QueryTable
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.QueryTable>(this, "QueryTable", typeof(NetOffice.ExcelApi.QueryTable));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839404.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Range Range
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Range", typeof(NetOffice.ExcelApi.Range));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837833.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual bool ShowAutoFilter
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
		public virtual bool ShowTotals
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
		public virtual NetOffice.ExcelApi.Enums.XlListObjectSourceType SourceType
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
		public virtual NetOffice.ExcelApi.Range TotalsRowRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "TotalsRowRange", typeof(NetOffice.ExcelApi.Range));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837420.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual string SharePointURL
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
		public virtual NetOffice.ExcelApi.XmlMap XmlMap
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.XmlMap>(this, "XmlMap", typeof(NetOffice.ExcelApi.XmlMap));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193006.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public virtual string DisplayName
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
		public virtual bool ShowHeaders
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
		public virtual NetOffice.ExcelApi.AutoFilter AutoFilter
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.AutoFilter>(this, "AutoFilter", typeof(NetOffice.ExcelApi.AutoFilter));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840453.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public virtual object TableStyle
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
		public virtual bool ShowTableStyleFirstColumn
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
		public virtual bool ShowTableStyleLastColumn
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
		public virtual bool ShowTableStyleRowStripes
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
		public virtual bool ShowTableStyleColumnStripes
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
		public virtual NetOffice.ExcelApi.Sort Sort
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sort>(this, "Sort", typeof(NetOffice.ExcelApi.Sort));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822164.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public virtual string Comment
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
		public virtual string AlternativeText
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
		public virtual string Summary
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
		public virtual NetOffice.ExcelApi.TableObject TableObject
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.TableObject>(this, "TableObject", typeof(NetOffice.ExcelApi.TableObject));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230953.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual NetOffice.ExcelApi.Slicers Slicers
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Slicers>(this, "Slicers", typeof(NetOffice.ExcelApi.Slicers));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231020.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual bool ShowAutoFilterDropDown
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
		public virtual void Delete()
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
		public virtual string Publish(object target, bool linkSource)
		{
			return Factory.ExecuteStringMethodGet(this, "Publish", target, linkSource);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834313.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual void Refresh()
		{
			 Factory.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196609.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual void Unlink()
		{
			 Factory.ExecuteMethod(this, "Unlink");
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193017.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual void Unlist()
		{
			 Factory.ExecuteMethod(this, "Unlist");
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="iConflictType">optional NetOffice.ExcelApi.Enums.XlListConflict iConflictType = 0</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual void UpdateChanges(object iConflictType)
		{
			 Factory.ExecuteMethod(this, "UpdateChanges", iConflictType);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual void UpdateChanges()
		{
			 Factory.ExecuteMethod(this, "UpdateChanges");
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838369.aspx </remarks>
		/// <param name="range">NetOffice.ExcelApi.Range range</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual void Resize(NetOffice.ExcelApi.Range range)
		{
			 Factory.ExecuteMethod(this, "Resize", range);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196053.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public virtual void ExportToVisio()
		{
			 Factory.ExecuteMethod(this, "ExportToVisio");
		}

		#endregion

		#pragma warning restore
	}
}


