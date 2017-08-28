using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface TableOfAuthorities 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836149.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class TableOfAuthorities : COMObject
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
                    _type = typeof(TableOfAuthorities);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public TableOfAuthorities(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public TableOfAuthorities(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableOfAuthorities(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableOfAuthorities(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableOfAuthorities(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableOfAuthorities(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableOfAuthorities() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TableOfAuthorities(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195937.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192349.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195036.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196811.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Passim
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Passim");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Passim", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840760.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool KeepEntryFormatting
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "KeepEntryFormatting");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "KeepEntryFormatting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192631.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Category
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Category");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Category", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820783.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string Bookmark
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Bookmark");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Bookmark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195391.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string Separator
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Separator");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Separator", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845894.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string IncludeSequenceName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "IncludeSequenceName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IncludeSequenceName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838521.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string EntrySeparator
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EntrySeparator");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EntrySeparator", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840432.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string PageRangeSeparator
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PageRangeSeparator");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageRangeSeparator", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197175.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool IncludeCategoryHeader
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IncludeCategoryHeader");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IncludeCategoryHeader", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839128.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string PageNumberSeparator
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PageNumberSeparator");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageNumberSeparator", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845490.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836913.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdTabLeader TabLeader
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdTabLeader>(this, "TabLeader");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TabLeader", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193018.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198214.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Update()
		{
			 Factory.ExecuteMethod(this, "Update");
		}

		#endregion

		#pragma warning restore
	}
}
