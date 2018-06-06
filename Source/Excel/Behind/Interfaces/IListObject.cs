using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface IListObject 
    /// SupportByVersion Excel, 11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IListObject : COMObject, NetOffice.ExcelApi.IListObject
    {
        #pragma warning disable

        #region Type Information

        /// <summary>        /// Instance Type
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
                    _type = typeof(IListObject);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IListObject() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16), ProxyResult]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 14, 15, 16)]
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
        [SupportByVersion("Excel", 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual Int32 Delete()
        {
            return Factory.ExecuteInt32MethodGet(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="target">object target</param>
        /// <param name="linkSource">bool linkSource</param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual string Publish(object target, bool linkSource)
        {
            return Factory.ExecuteStringMethodGet(this, "Publish", target, linkSource);
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual Int32 Refresh()
        {
            return Factory.ExecuteInt32MethodGet(this, "Refresh");
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual Int32 Unlink()
        {
            return Factory.ExecuteInt32MethodGet(this, "Unlink");
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual Int32 Unlist()
        {
            return Factory.ExecuteInt32MethodGet(this, "Unlist");
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="iConflictType">optional NetOffice.ExcelApi.Enums.XlListConflict iConflictType = 0</param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual Int32 UpdateChanges(object iConflictType)
        {
            return Factory.ExecuteInt32MethodGet(this, "UpdateChanges", iConflictType);
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual Int32 UpdateChanges()
        {
            return Factory.ExecuteInt32MethodGet(this, "UpdateChanges");
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="range">NetOffice.ExcelApi.Range range</param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual Int32 Resize(NetOffice.ExcelApi.Range range)
        {
            return Factory.ExecuteInt32MethodGet(this, "Resize", range);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportToVisio()
        {
            return Factory.ExecuteInt32MethodGet(this, "ExportToVisio");
        }

        #endregion

        #pragma warning restore
    }
}

