using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Table 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834860.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Table : COMObject
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
                    _type = typeof(Table);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Table(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Table(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197195.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Range
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Range", NetOffice.WordApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839082.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845545.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835743.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198160.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Columns Columns
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Columns>(this, "Columns", NetOffice.WordApi.Columns.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839587.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Rows Rows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Rows>(this, "Rows", NetOffice.WordApi.Rows.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823239.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Borders Borders
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Borders>(this, "Borders", NetOffice.WordApi.Borders.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Borders", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845404.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shading Shading
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Shading>(this, "Shading", NetOffice.WordApi.Shading.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835471.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Uniform
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Uniform");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193447.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 AutoFormatType
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AutoFormatType");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836124.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Tables Tables
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Tables>(this, "Tables", NetOffice.WordApi.Tables.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194409.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 NestingLevel
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "NestingLevel");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AllowPageBreaks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowPageBreaks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowPageBreaks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839810.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool AllowAutoFit
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowAutoFit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowAutoFit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845887.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single PreferredWidth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "PreferredWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PreferredWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834288.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdPreferredWidthType PreferredWidthType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdPreferredWidthType>(this, "PreferredWidthType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PreferredWidthType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844783.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single TopPadding
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "TopPadding");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TopPadding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838742.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single BottomPadding
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BottomPadding");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BottomPadding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836311.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single LeftPadding
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "LeftPadding");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LeftPadding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835709.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single RightPadding
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "RightPadding");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RightPadding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196121.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single Spacing
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Spacing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Spacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193098.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdTableDirection TableDirection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdTableDirection>(this, "TableDirection");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TableDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840350.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string ID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ID");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ID", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196552.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
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
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191959.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool ApplyStyleHeadingRows
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ApplyStyleHeadingRows");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ApplyStyleHeadingRows", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844792.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool ApplyStyleLastRow
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ApplyStyleLastRow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ApplyStyleLastRow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834832.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool ApplyStyleFirstColumn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ApplyStyleFirstColumn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ApplyStyleFirstColumn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839138.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool ApplyStyleLastColumn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ApplyStyleLastColumn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ApplyStyleLastColumn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192619.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool ApplyStyleRowBands
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ApplyStyleRowBands");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ApplyStyleRowBands", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839106.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool ApplyStyleColumnBands
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ApplyStyleColumnBands");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ApplyStyleColumnBands", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835972.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public string Title
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Title");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Title", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820918.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public string Descr
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Descr");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Descr", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194359.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845868.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="languageID">optional object languageID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object languageID)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, languageID });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld()
		{
			 Factory.ExecuteMethod(this, "SortOld");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader)
		{
			 Factory.ExecuteMethod(this, "SortOld", excludeHeader);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber)
		{
			 Factory.ExecuteMethod(this, "SortOld", excludeHeader, fieldNumber);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType)
		{
			 Factory.ExecuteMethod(this, "SortOld", excludeHeader, fieldNumber, sortFieldType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder)
		{
			 Factory.ExecuteMethod(this, "SortOld", excludeHeader, fieldNumber, sortFieldType, sortOrder);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196507.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortAscending()
		{
			 Factory.ExecuteMethod(this, "SortAscending");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835818.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortDescending()
		{
			 Factory.ExecuteMethod(this, "SortDescending");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		/// <param name="applyLastRow">optional object applyLastRow</param>
		/// <param name="applyFirstColumn">optional object applyFirstColumn</param>
		/// <param name="applyLastColumn">optional object applyLastColumn</param>
		/// <param name="autoFit">optional object autoFit</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit)
		{
			 Factory.ExecuteMethod(this, "AutoFormat", new object[]{ format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat()
		{
			 Factory.ExecuteMethod(this, "AutoFormat");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format)
		{
			 Factory.ExecuteMethod(this, "AutoFormat", format);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format, object applyBorders)
		{
			 Factory.ExecuteMethod(this, "AutoFormat", format, applyBorders);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format, object applyBorders, object applyShading)
		{
			 Factory.ExecuteMethod(this, "AutoFormat", format, applyBorders, applyShading);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format, object applyBorders, object applyShading, object applyFont)
		{
			 Factory.ExecuteMethod(this, "AutoFormat", format, applyBorders, applyShading, applyFont);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor)
		{
			 Factory.ExecuteMethod(this, "AutoFormat", new object[]{ format, applyBorders, applyShading, applyFont, applyColor });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows)
		{
			 Factory.ExecuteMethod(this, "AutoFormat", new object[]{ format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		/// <param name="applyLastRow">optional object applyLastRow</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow)
		{
			 Factory.ExecuteMethod(this, "AutoFormat", new object[]{ format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		/// <param name="applyLastRow">optional object applyLastRow</param>
		/// <param name="applyFirstColumn">optional object applyFirstColumn</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn)
		{
			 Factory.ExecuteMethod(this, "AutoFormat", new object[]{ format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		/// <param name="applyLastRow">optional object applyLastRow</param>
		/// <param name="applyFirstColumn">optional object applyFirstColumn</param>
		/// <param name="applyLastColumn">optional object applyLastColumn</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn)
		{
			 Factory.ExecuteMethod(this, "AutoFormat", new object[]{ format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838712.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void UpdateAutoFormat()
		{
			 Factory.ExecuteMethod(this, "UpdateAutoFormat");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="separator">optional object separator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range ConvertToTextOld(object separator)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "ConvertToTextOld", NetOffice.WordApi.Range.LateBindingApiWrapperType, separator);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range ConvertToTextOld()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "ConvertToTextOld", NetOffice.WordApi.Range.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821612.aspx </remarks>
		/// <param name="row">Int32 row</param>
		/// <param name="column">Int32 column</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Cell Cell(Int32 row, Int32 column)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Cell>(this, "Cell", NetOffice.WordApi.Cell.LateBindingApiWrapperType, row, column);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836035.aspx </remarks>
		/// <param name="beforeRow">object beforeRow</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table Split(object beforeRow)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "Split", NetOffice.WordApi.Table.LateBindingApiWrapperType, beforeRow);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820974.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		/// <param name="nestedTables">optional object nestedTables</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range ConvertToText(object separator, object nestedTables)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "ConvertToText", NetOffice.WordApi.Range.LateBindingApiWrapperType, separator, nestedTables);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820974.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range ConvertToText()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "ConvertToText", NetOffice.WordApi.Range.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820974.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range ConvertToText(object separator)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "ConvertToText", NetOffice.WordApi.Range.LateBindingApiWrapperType, separator);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820953.aspx </remarks>
		/// <param name="behavior">NetOffice.WordApi.Enums.WdAutoFitBehavior behavior</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AutoFitBehavior(NetOffice.WordApi.Enums.WdAutoFitBehavior behavior)
		{
			 Factory.ExecuteMethod(this, "AutoFitBehavior", behavior);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		/// <param name="ignoreHe">optional object ignoreHe</param>
		/// <param name="languageID">optional object languageID</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort()
		{
			 Factory.ExecuteMethod(this, "Sort");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader)
		{
			 Factory.ExecuteMethod(this, "Sort", excludeHeader);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber)
		{
			 Factory.ExecuteMethod(this, "Sort", excludeHeader, fieldNumber);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType)
		{
			 Factory.ExecuteMethod(this, "Sort", excludeHeader, fieldNumber, sortFieldType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder)
		{
			 Factory.ExecuteMethod(this, "Sort", excludeHeader, fieldNumber, sortFieldType, sortOrder);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, bidiSort });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, bidiSort, ignoreThe });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, bidiSort, ignoreThe, ignoreKashida });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		/// <param name="ignoreHe">optional object ignoreHe</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192363.aspx </remarks>
		/// <param name="styleName">string styleName</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ApplyStyleDirectFormatting(string styleName)
		{
			 Factory.ExecuteMethod(this, "ApplyStyleDirectFormatting", styleName);
		}

		#endregion

		#pragma warning restore
	}
}
