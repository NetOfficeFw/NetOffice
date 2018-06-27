using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface Table 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834860.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Table : COMObject, NetOffice.WordApi.Table
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
                    _contractType = typeof(NetOffice.WordApi.Table);
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
                    _type = typeof(Table);                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Table() : base()
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
		public virtual NetOffice.WordApi.Range Range
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Range", typeof(NetOffice.WordApi.Range));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839082.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845545.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835743.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198160.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Columns Columns
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Columns>(this, "Columns", typeof(NetOffice.WordApi.Columns));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839587.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Rows Rows
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Rows>(this, "Rows", typeof(NetOffice.WordApi.Rows));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823239.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Borders Borders
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Borders>(this, "Borders", typeof(NetOffice.WordApi.Borders));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Borders", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845404.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Shading Shading
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Shading>(this, "Shading", typeof(NetOffice.WordApi.Shading));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835471.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool Uniform
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Uniform");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193447.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 AutoFormatType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "AutoFormatType");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836124.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Tables Tables
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Tables>(this, "Tables", typeof(NetOffice.WordApi.Tables));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194409.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 NestingLevel
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "NestingLevel");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool AllowPageBreaks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowPageBreaks");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowPageBreaks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839810.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool AllowAutoFit
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowAutoFit");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowAutoFit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845887.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single PreferredWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "PreferredWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PreferredWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834288.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdPreferredWidthType PreferredWidthType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdPreferredWidthType>(this, "PreferredWidthType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PreferredWidthType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844783.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single TopPadding
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "TopPadding");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TopPadding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838742.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single BottomPadding
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "BottomPadding");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BottomPadding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836311.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single LeftPadding
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "LeftPadding");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LeftPadding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835709.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single RightPadding
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "RightPadding");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RightPadding", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196121.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single Spacing
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "Spacing");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Spacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193098.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdTableDirection TableDirection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdTableDirection>(this, "TableDirection");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "TableDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840350.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string ID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ID");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ID", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196552.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual object Style
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Style");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Style", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191959.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool ApplyStyleHeadingRows
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ApplyStyleHeadingRows");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ApplyStyleHeadingRows", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844792.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool ApplyStyleLastRow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ApplyStyleLastRow");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ApplyStyleLastRow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834832.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool ApplyStyleFirstColumn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ApplyStyleFirstColumn");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ApplyStyleFirstColumn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839138.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool ApplyStyleLastColumn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ApplyStyleLastColumn");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ApplyStyleLastColumn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192619.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool ApplyStyleRowBands
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ApplyStyleRowBands");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ApplyStyleRowBands", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839106.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool ApplyStyleColumnBands
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ApplyStyleColumnBands");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ApplyStyleColumnBands", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835972.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual string Title
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Title");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Title", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820918.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual string Descr
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Descr");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Descr", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194359.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Select()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845868.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Delete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
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
		public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object languageID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, languageID });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SortOld()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SortOld(object excludeHeader)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", excludeHeader);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SortOld(object excludeHeader, object fieldNumber)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", excludeHeader, fieldNumber);
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
		public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", excludeHeader, fieldNumber, sortFieldType);
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
		public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", excludeHeader, fieldNumber, sortFieldType, sortOrder);
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
		public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2 });
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
		public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2 });
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
		public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2 });
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
		public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3 });
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
		public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3 });
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
		public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3 });
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
		public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196507.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SortAscending()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SortAscending");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835818.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SortDescending()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SortDescending");
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
		public virtual void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat", new object[]{ format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void AutoFormat()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void AutoFormat(object format)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat", format);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void AutoFormat(object format, object applyBorders)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat", format, applyBorders);
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
		public virtual void AutoFormat(object format, object applyBorders, object applyShading)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat", format, applyBorders, applyShading);
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
		public virtual void AutoFormat(object format, object applyBorders, object applyShading, object applyFont)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat", format, applyBorders, applyShading, applyFont);
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
		public virtual void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat", new object[]{ format, applyBorders, applyShading, applyFont, applyColor });
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
		public virtual void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat", new object[]{ format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows });
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
		public virtual void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat", new object[]{ format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow });
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
		public virtual void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat", new object[]{ format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn });
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
		public virtual void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat", new object[]{ format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838712.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void UpdateAutoFormat()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateAutoFormat");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="separator">optional object separator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range ConvertToTextOld(object separator)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "ConvertToTextOld", typeof(NetOffice.WordApi.Range), separator);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range ConvertToTextOld()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "ConvertToTextOld", typeof(NetOffice.WordApi.Range));
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821612.aspx </remarks>
		/// <param name="row">Int32 row</param>
		/// <param name="column">Int32 column</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Cell Cell(Int32 row, Int32 column)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Cell>(this, "Cell", typeof(NetOffice.WordApi.Cell), row, column);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836035.aspx </remarks>
		/// <param name="beforeRow">object beforeRow</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Table Split(object beforeRow)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "Split", typeof(NetOffice.WordApi.Table), beforeRow);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820974.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		/// <param name="nestedTables">optional object nestedTables</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range ConvertToText(object separator, object nestedTables)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "ConvertToText", typeof(NetOffice.WordApi.Range), separator, nestedTables);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820974.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range ConvertToText()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "ConvertToText", typeof(NetOffice.WordApi.Range));
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820974.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range ConvertToText(object separator)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "ConvertToText", typeof(NetOffice.WordApi.Range), separator);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820953.aspx </remarks>
		/// <param name="behavior">NetOffice.WordApi.Enums.WdAutoFitBehavior behavior</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void AutoFitBehavior(NetOffice.WordApi.Enums.WdAutoFitBehavior behavior)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFitBehavior", behavior);
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
		public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Sort()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Sort");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Sort(object excludeHeader)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", excludeHeader);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Sort(object excludeHeader, object fieldNumber)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", excludeHeader, fieldNumber);
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
		public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", excludeHeader, fieldNumber, sortFieldType);
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
		public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", excludeHeader, fieldNumber, sortFieldType, sortOrder);
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
		public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2 });
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
		public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2 });
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
		public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2 });
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
		public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3 });
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
		public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3 });
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
		public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3 });
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
		public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive });
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
		public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, bidiSort });
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
		public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, bidiSort, ignoreThe });
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
		public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, bidiSort, ignoreThe, ignoreKashida });
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
		public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics });
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
		public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192363.aspx </remarks>
		/// <param name="styleName">string styleName</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ApplyStyleDirectFormatting(string styleName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyStyleDirectFormatting", styleName);
		}

		#endregion

		#pragma warning restore
	}
}


