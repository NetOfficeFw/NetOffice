using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface _QueryTable 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _QueryTable : COMObject
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
                    _type = typeof(_QueryTable);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _QueryTable(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _QueryTable(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _QueryTable(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _QueryTable(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _QueryTable(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _QueryTable(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _QueryTable() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _QueryTable(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.Parent"/> </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.Name"/> </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.FieldNames"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool FieldNames
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FieldNames");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FieldNames", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.RowNumbers"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool RowNumbers
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RowNumbers");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RowNumbers", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.FillAdjacentFormulas"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool FillAdjacentFormulas
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FillAdjacentFormulas");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FillAdjacentFormulas", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool HasAutoFormat
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasAutoFormat");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasAutoFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.RefreshOnFileOpen"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool RefreshOnFileOpen
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RefreshOnFileOpen");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RefreshOnFileOpen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.Refreshing"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Refreshing
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Refreshing");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.FetchedRowOverflow"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool FetchedRowOverflow
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FetchedRowOverflow");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.BackgroundQuery"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool BackgroundQuery
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BackgroundQuery");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackgroundQuery", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.RefreshStyle"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCellInsertionMode RefreshStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCellInsertionMode>(this, "RefreshStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RefreshStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.EnableRefresh"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool EnableRefresh
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableRefresh");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableRefresh", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.SavePassword"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool SavePassword
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SavePassword");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SavePassword", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.Destination"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Destination
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Destination", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.Connection"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Connection
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Connection");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Connection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Sql
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Sql");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Sql", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.PostText"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string PostText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PostText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PostText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.ResultRange"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range ResultRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "ResultRange", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.Parameters"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Parameters Parameters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Parameters>(this, "Parameters", NetOffice.ExcelApi.Parameters.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.Recordset"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object Recordset
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Recordset");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Recordset", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.SaveData"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool SaveData
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SaveData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SaveData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool TablesOnlyFromHTML
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TablesOnlyFromHTML");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TablesOnlyFromHTML", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.EnableEditing"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool EnableEditing
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableEditing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableEditing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.TextFilePlatform"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlPlatform TextFilePlatform
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPlatform>(this, "TextFilePlatform");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TextFilePlatform", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.TextFileStartRow"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 TextFileStartRow
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TextFileStartRow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileStartRow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.TextFileParseType"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlTextParsingType TextFileParseType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTextParsingType>(this, "TextFileParseType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TextFileParseType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.TextFileTextQualifier"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlTextQualifier TextFileTextQualifier
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTextQualifier>(this, "TextFileTextQualifier");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TextFileTextQualifier", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.TextFileConsecutiveDelimiter"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool TextFileConsecutiveDelimiter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TextFileConsecutiveDelimiter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileConsecutiveDelimiter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.TextFileTabDelimiter"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool TextFileTabDelimiter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TextFileTabDelimiter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileTabDelimiter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.TextFileSemicolonDelimiter"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool TextFileSemicolonDelimiter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TextFileSemicolonDelimiter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileSemicolonDelimiter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.TextFileCommaDelimiter"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool TextFileCommaDelimiter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TextFileCommaDelimiter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileCommaDelimiter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.TextFileSpaceDelimiter"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool TextFileSpaceDelimiter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TextFileSpaceDelimiter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileSpaceDelimiter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.TextFileOtherDelimiter"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string TextFileOtherDelimiter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TextFileOtherDelimiter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileOtherDelimiter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.TextFileColumnDataTypes"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TextFileColumnDataTypes
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TextFileColumnDataTypes");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "TextFileColumnDataTypes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.TextFileFixedColumnWidths"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object TextFileFixedColumnWidths
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TextFileFixedColumnWidths");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "TextFileFixedColumnWidths", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.PreserveColumnInfo"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool PreserveColumnInfo
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PreserveColumnInfo");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PreserveColumnInfo", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.PreserveFormatting"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool PreserveFormatting
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PreserveFormatting");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PreserveFormatting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.AdjustColumnWidth"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool AdjustColumnWidth
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AdjustColumnWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AdjustColumnWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.CommandText"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object CommandText
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CommandText");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "CommandText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.CommandType"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCmdType CommandType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCmdType>(this, "CommandType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CommandType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.TextFilePromptOnRefresh"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool TextFilePromptOnRefresh
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TextFilePromptOnRefresh");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFilePromptOnRefresh", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.QueryType"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.xlQueryType QueryType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.xlQueryType>(this, "QueryType");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.MaintainConnection"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool MaintainConnection
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MaintainConnection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaintainConnection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.TextFileDecimalSeparator"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string TextFileDecimalSeparator
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TextFileDecimalSeparator");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileDecimalSeparator", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.TextFileThousandsSeparator"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string TextFileThousandsSeparator
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TextFileThousandsSeparator");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileThousandsSeparator", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.RefreshPeriod"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 RefreshPeriod
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RefreshPeriod");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RefreshPeriod", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.WebSelectionType"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.xlWebSelectionType WebSelectionType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.xlWebSelectionType>(this, "WebSelectionType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "WebSelectionType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.WebFormatting"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.xlWebFormatting WebFormatting
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.xlWebFormatting>(this, "WebFormatting");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "WebFormatting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.WebTables"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string WebTables
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "WebTables");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WebTables", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.WebPreFormattedTextToColumns"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool WebPreFormattedTextToColumns
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WebPreFormattedTextToColumns");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WebPreFormattedTextToColumns", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.WebSingleBlockTextImport"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool WebSingleBlockTextImport
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WebSingleBlockTextImport");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WebSingleBlockTextImport", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.WebDisableDateRecognition"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool WebDisableDateRecognition
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WebDisableDateRecognition");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WebDisableDateRecognition", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.WebConsecutiveDelimitersAsOne"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool WebConsecutiveDelimitersAsOne
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WebConsecutiveDelimitersAsOne");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WebConsecutiveDelimitersAsOne", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.WebDisableRedirections"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool WebDisableRedirections
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WebDisableRedirections");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WebDisableRedirections", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.EditWebPage"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public object EditWebPage
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EditWebPage");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EditWebPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.SourceConnectionFile"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public string SourceConnectionFile
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SourceConnectionFile");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SourceConnectionFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.SourceDataFile"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public string SourceDataFile
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SourceDataFile");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SourceDataFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.RobustConnect"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlRobustConnect RobustConnect
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlRobustConnect>(this, "RobustConnect");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RobustConnect", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.TextFileTrailingMinusNumbers"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool TextFileTrailingMinusNumbers
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TextFileTrailingMinusNumbers");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileTrailingMinusNumbers", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.ListObject"/> </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.ListObject ListObject
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ListObject>(this, "ListObject", NetOffice.ExcelApi.ListObject.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.TextFileVisualLayout"/> </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlTextVisualLayoutType TextFileVisualLayout
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTextVisualLayoutType>(this, "TextFileVisualLayout");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TextFileVisualLayout", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.WorkbookConnection"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.WorkbookConnection WorkbookConnection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.WorkbookConnection>(this, "WorkbookConnection", NetOffice.ExcelApi.WorkbookConnection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.Sort"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Sort Sort
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sort>(this, "Sort", NetOffice.ExcelApi.Sort.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.CancelRefresh"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void CancelRefresh()
		{
			 Factory.ExecuteMethod(this, "CancelRefresh");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.Delete"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.Refresh"/> </remarks>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Refresh(object backgroundQuery)
		{
			return Factory.ExecuteBoolMethodGet(this, "Refresh", backgroundQuery);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.Refresh"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Refresh()
		{
			return Factory.ExecuteBoolMethodGet(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.ResetTimer"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void ResetTimer()
		{
			 Factory.ExecuteMethod(this, "ResetTimer");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.SaveAsODC"/> </remarks>
		/// <param name="oDCFileName">string oDCFileName</param>
		/// <param name="description">optional object description</param>
		/// <param name="keywords">optional object keywords</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void SaveAsODC(string oDCFileName, object description, object keywords)
		{
			 Factory.ExecuteMethod(this, "SaveAsODC", oDCFileName, description, keywords);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.SaveAsODC"/> </remarks>
		/// <param name="oDCFileName">string oDCFileName</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void SaveAsODC(string oDCFileName)
		{
			 Factory.ExecuteMethod(this, "SaveAsODC", oDCFileName);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.QueryTable.SaveAsODC"/> </remarks>
		/// <param name="oDCFileName">string oDCFileName</param>
		/// <param name="description">optional object description</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void SaveAsODC(string oDCFileName, object description)
		{
			 Factory.ExecuteMethod(this, "SaveAsODC", oDCFileName, description);
		}

		#endregion

		#pragma warning restore
	}
}
