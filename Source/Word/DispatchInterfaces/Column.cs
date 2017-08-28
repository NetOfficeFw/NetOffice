using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Column 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195142.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Column : COMObject
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
                    _type = typeof(Column);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Column(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Column(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Column(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Column(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Column(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Column(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Column() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Column(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192790.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838273.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837510.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820964.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single Width
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Width");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194360.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool IsFirst
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsFirst");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835237.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool IsLast
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsLast");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840774.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Index
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194516.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Cells Cells
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Cells>(this, "Cells", NetOffice.WordApi.Cells.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840574.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838949.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840814.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Column Next
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Column>(this, "Next", NetOffice.WordApi.Column.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197187.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Column Previous
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Column>(this, "Previous", NetOffice.WordApi.Column.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191824.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836716.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193135.aspx </remarks>
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193096.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192019.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840911.aspx </remarks>
		/// <param name="columnWidth">Single columnWidth</param>
		/// <param name="rulerStyle">NetOffice.WordApi.Enums.WdRulerStyle rulerStyle</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SetWidth(Single columnWidth, NetOffice.WordApi.Enums.WdRulerStyle rulerStyle)
		{
			 Factory.ExecuteMethod(this, "SetWidth", columnWidth, rulerStyle);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838466.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AutoFit()
		{
			 Factory.ExecuteMethod(this, "AutoFit");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="languageID">optional object languageID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object sortFieldType, object sortOrder, object caseSensitive, object languageID)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, sortFieldType, sortOrder, caseSensitive, languageID });
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
		/// <param name="sortFieldType">optional object sortFieldType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object sortFieldType)
		{
			 Factory.ExecuteMethod(this, "SortOld", excludeHeader, sortFieldType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object sortFieldType, object sortOrder)
		{
			 Factory.ExecuteMethod(this, "SortOld", excludeHeader, sortFieldType, sortOrder);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object sortFieldType, object sortOrder, object caseSensitive)
		{
			 Factory.ExecuteMethod(this, "SortOld", excludeHeader, sortFieldType, sortOrder, caseSensitive);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838077.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		/// <param name="ignoreHe">optional object ignoreHe</param>
		/// <param name="languageID">optional object languageID</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838077.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort()
		{
			 Factory.ExecuteMethod(this, "Sort");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838077.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838077.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object sortFieldType)
		{
			 Factory.ExecuteMethod(this, "Sort", excludeHeader, sortFieldType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838077.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object sortFieldType, object sortOrder)
		{
			 Factory.ExecuteMethod(this, "Sort", excludeHeader, sortFieldType, sortOrder);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838077.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object sortFieldType, object sortOrder, object caseSensitive)
		{
			 Factory.ExecuteMethod(this, "Sort", excludeHeader, sortFieldType, sortOrder, caseSensitive);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838077.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object sortFieldType, object sortOrder, object caseSensitive, object bidiSort)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, sortFieldType, sortOrder, caseSensitive, bidiSort });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838077.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838077.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838077.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838077.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		/// <param name="ignoreHe">optional object ignoreHe</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe });
		}

		#endregion

		#pragma warning restore
	}
}
