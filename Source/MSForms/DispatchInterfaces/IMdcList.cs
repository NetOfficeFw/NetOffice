using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// IMdcList
	/// </summary>
	[SyntaxBypass]
 	public class IMdcList_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IMdcList_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IMdcList_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcList_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcList_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcList_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcList_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcList_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcList_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		/// <param name="pvargColumn">optional object pvargColumn</param>
		/// <param name="pvargIndex">optional object pvargIndex</param>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_Column(object pvargColumn, object pvargIndex)
		{
			return Factory.ExecuteVariantPropertyGet(this, "Column", pvargColumn, pvargIndex);
		}

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        /// <param name="pvargColumn">optional object pvargColumn</param>
        /// <param name="pvargIndex">optional object pvargIndex</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Column(object pvargColumn, object pvargIndex, object value)
		{
			Factory.ExecutePropertySet(this, "Column", pvargColumn, pvargIndex, value);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Alias for get_Column
		/// </summary>
		/// <param name="pvargColumn">optional object pvargColumn</param>
		/// <param name="pvargIndex">optional object pvargIndex</param>
		[SupportByVersion("MSForms", 2), Redirect("get_Column")]
		public object Column(object pvargColumn, object pvargIndex)
		{
			return get_Column(pvargColumn, pvargIndex);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		/// <param name="pvargColumn">optional object pvargColumn</param>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_Column(object pvargColumn)
		{
			return Factory.ExecuteVariantPropertyGet(this, "Column", pvargColumn);
		}

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        /// <param name="pvargColumn">optional object pvargColumn</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Column(object pvargColumn, object value)
		{
			Factory.ExecutePropertySet(this, "Column", pvargColumn, value);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Alias for get_Column
		/// </summary>
		/// <param name="pvargColumn">optional object pvargColumn</param>
		[SupportByVersion("MSForms", 2), Redirect("get_Column")]
		public object Column(object pvargColumn)
		{
			return get_Column(pvargColumn);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		/// <param name="pvargIndex">optional object pvargIndex</param>
		/// <param name="pvargColumn">optional object pvargColumn</param>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_List(object pvargIndex, object pvargColumn)
		{
			return Factory.ExecuteVariantPropertyGet(this, "List", pvargIndex, pvargColumn);
		}

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        /// <param name="pvargIndex">optional object pvargIndex</param>
        /// <param name="pvargColumn">optional object pvargColumn</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_List(object pvargIndex, object pvargColumn, object value)
		{
			Factory.ExecutePropertySet(this, "List", pvargIndex, pvargColumn, value);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Alias for get_List
		/// </summary>
		/// <param name="pvargIndex">optional object pvargIndex</param>
		/// <param name="pvargColumn">optional object pvargColumn</param>
		[SupportByVersion("MSForms", 2), Redirect("get_List")]
		public object List(object pvargIndex, object pvargColumn)
		{
			return get_List(pvargIndex, pvargColumn);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		/// <param name="pvargIndex">optional object pvargIndex</param>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_List(object pvargIndex)
		{
			return Factory.ExecuteVariantPropertyGet(this, "List", pvargIndex);
		}

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        /// <param name="pvargIndex">optional object pvargIndex</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_List(object pvargIndex, object value)
		{
			Factory.ExecutePropertySet(this, "List", pvargIndex, value);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Alias for get_List
		/// </summary>
		/// <param name="pvargIndex">optional object pvargIndex</param>
		[SupportByVersion("MSForms", 2), Redirect("get_List")]
		public object List(object pvargIndex)
		{
			return get_List(pvargIndex);
		}

		#endregion

		#region Methods

		#endregion

	}

	/// <summary>
	/// DispatchInterface IMdcList 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IMdcList : IMdcList_
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
                    _type = typeof(IMdcList);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IMdcList(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IMdcList(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcList(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcList(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcList(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcList(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcList() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcList(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int32 BackColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BackColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int32 BorderColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BorderColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BorderColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmBorderStyle BorderStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmBorderStyle>(this, "BorderStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BorderStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool BordersSuppress
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BordersSuppress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BordersSuppress", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public object BoundColumn
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BoundColumn");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BoundColumn", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int32 ColumnCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ColumnCount");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ColumnCount", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool ColumnHeads
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ColumnHeads");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ColumnHeads", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public string ColumnWidths
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ColumnWidths");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ColumnWidths", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool Enabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Enabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Enabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Font _Font_Reserved
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSFormsApi.Font>(this, "_Font_Reserved");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "_Font_Reserved", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[BaseResult]
		public NetOffice.MSFormsApi.Font Font
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSFormsApi.Font>(this, "Font");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Font", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool FontBold
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FontBold");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontBold", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool FontItalic
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FontItalic");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontItalic", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string FontName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FontName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontName", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public float FontSize
		{
			get
			{
				return Factory.ExecuteFloatPropertyGet(this, "FontSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool FontStrikethru
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FontStrikethru");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontStrikethru", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool FontUnderline
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FontUnderline");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontUnderline", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 FontWeight
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "FontWeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontWeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int32 ForeColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ForeColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool IntegralHeight
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IntegralHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IntegralHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 ListCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ListCount");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSForms", 2), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ListCursor
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ListCursor");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "ListCursor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ListIndex
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ListIndex");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ListIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmListStyle ListStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmListStyle>(this, "ListStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ListStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ListWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ListWidth");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ListWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
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
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmMatchEntry MatchEntry
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmMatchEntry>(this, "MatchEntry");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MatchEntry", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2), NativeResult]
		public stdole.Picture MouseIcon
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MouseIcon", paramsArray);
                return returnItem as stdole.Picture;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MouseIcon", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmMousePointer MousePointer
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmMousePointer>(this, "MousePointer");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MousePointer", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmMultiSelect MultiSelect
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmMultiSelect>(this, "MultiSelect");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MultiSelect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmSpecialEffect SpecialEffect
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmSpecialEffect>(this, "SpecialEffect");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SpecialEffect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public string Text
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public object TextColumn
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TextColumn");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "TextColumn", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public object TopIndex
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TopIndex");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "TopIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool Valid
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Valid");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public object Value
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Value");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Value", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Column
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Column");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Column", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object List
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "List");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "List", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		/// <param name="pvargIndex">object pvargIndex</param>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool get_Selected(object pvargIndex)
		{
			return Factory.ExecuteBoolPropertyGet(this, "Selected", pvargIndex);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		/// <param name="pvargIndex">object pvargIndex</param>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Selected(object pvargIndex, bool value)
		{
			Factory.ExecutePropertySet(this, "Selected", pvargIndex, value);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Alias for get_Selected
		/// </summary>
		/// <param name="pvargIndex">object pvargIndex</param>
		[SupportByVersion("MSForms", 2), Redirect("get_Selected")]
		public bool Selected(object pvargIndex)
		{
			return get_Selected(pvargIndex);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmIMEMode IMEMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmIMEMode>(this, "IMEMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "IMEMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Enums.fmDisplayStyle DisplayStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmDisplayStyle>(this, "DisplayStyle");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmTextAlign TextAlign
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmTextAlign>(this, "TextAlign");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TextAlign", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="pvargItem">optional object pvargItem</param>
		/// <param name="pvargIndex">optional object pvargIndex</param>
		[SupportByVersion("MSForms", 2)]
		public void AddItem(object pvargItem, object pvargIndex)
		{
			 Factory.ExecuteMethod(this, "AddItem", pvargItem, pvargIndex);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public void AddItem()
		{
			 Factory.ExecuteMethod(this, "AddItem");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="pvargItem">optional object pvargItem</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public void AddItem(object pvargItem)
		{
			 Factory.ExecuteMethod(this, "AddItem", pvargItem);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void Clear()
		{
			 Factory.ExecuteMethod(this, "Clear");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="pvargIndex">object pvargIndex</param>
		[SupportByVersion("MSForms", 2)]
		public void RemoveItem(object pvargIndex)
		{
			 Factory.ExecuteMethod(this, "RemoveItem", pvargIndex);
		}

		#endregion

		#pragma warning restore
	}
}
