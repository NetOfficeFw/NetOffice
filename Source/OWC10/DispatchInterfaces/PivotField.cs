using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotField 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PivotField : COMObject
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
                    _type = typeof(PivotField);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PivotField(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PivotField(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotField(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotField(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotField(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotField(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotField() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotField(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
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
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string BaseName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BaseName");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.ADODBApi.Enums.DataTypeEnum DataType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.DataTypeEnum>(this, "DataType");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
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
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 DetailWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DetailWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DetailWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 GroupedWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "GroupedWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GroupedWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		/// <param name="subtotals">Int32 subtotals</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool get_Subtotals(Int32 subtotals)
		{
			return Factory.ExecuteBoolPropertyGet(this, "Subtotals", subtotals);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		/// <param name="subtotals">Int32 subtotals</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Subtotals(Int32 subtotals, bool value)
		{
			Factory.ExecutePropertySet(this, "Subtotals", subtotals, value);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Subtotals
		/// </summary>
		/// <param name="subtotals">Int32 subtotals</param>
		[SupportByVersion("OWC10", 1), Redirect("get_Subtotals")]
		public bool Subtotals(Int32 subtotals)
		{
			return get_Subtotals(subtotals);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotFont DetailFont
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFont>(this, "DetailFont", NetOffice.OWC10Api.PivotFont.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object DetailForeColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DetailForeColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DetailForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object DetailBackColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DetailBackColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DetailBackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotHAlignmentEnum DetailHAlignment
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotHAlignmentEnum>(this, "DetailHAlignment");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DetailHAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotFont SubtotalFont
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFont>(this, "SubtotalFont", NetOffice.OWC10Api.PivotFont.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object SubtotalForeColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "SubtotalForeColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "SubtotalForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object SubtotalBackColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "SubtotalBackColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "SubtotalBackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotFieldGroupOnEnum GroupOn
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotFieldGroupOnEnum>(this, "GroupOn");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "GroupOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double GroupInterval
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "GroupInterval");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GroupInterval", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string Expression
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Expression");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Expression", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string NumberFormat
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NumberFormat");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NumberFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string DataField
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataField");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool IsIncluded
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsIncluded");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IsIncluded", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotFieldSortDirectionEnum SortDirection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotFieldSortDirectionEnum>(this, "SortDirection");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SortDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object OrderedMembers
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "OrderedMembers");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "OrderedMembers", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object MemberCaptions
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "MemberCaptions");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "MemberCaptions", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotFieldTypeEnum Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotFieldTypeEnum>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotFieldFilterFunctionEnum FilterFunction
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotFieldFilterFunctionEnum>(this, "FilterFunction");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FilterFunction", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DetailAutoFit
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DetailAutoFit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DetailAutoFit", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool GroupedAutoFit
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "GroupedAutoFit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GroupedAutoFit", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotFieldSet FieldSet
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFieldSet>(this, "FieldSet", NetOffice.OWC10Api.PivotFieldSet.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool Expanded
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Expanded");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Expanded", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotTotal SortOn
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotTotal>(this, "SortOn", NetOffice.OWC10Api.PivotTotal.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "SortOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object SortOnScope
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "SortOnScope");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "SortOnScope", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool IsHyperlink
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsHyperlink");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IsHyperlink", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string UniqueName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UniqueName");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object GroupStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "GroupStart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "GroupStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object GroupEnd
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "GroupEnd");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "GroupEnd", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object IncludedMembers
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "IncludedMembers");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "IncludedMembers", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object ExcludedMembers
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ExcludedMembers");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ExcludedMembers", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotMemberProperties MemberProperties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotMemberProperties>(this, "MemberProperties", NetOffice.OWC10Api.PivotMemberProperties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object MemberPropertiesOrder
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "MemberPropertiesOrder");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "MemberPropertiesOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 PropertyCaptionWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PropertyCaptionWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PropertyCaptionWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 GroupedHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "GroupedHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GroupedHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 PropertyValueWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PropertyValueWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PropertyValueWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 PropertyHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PropertyHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PropertyHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotField FilterContext
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotField>(this, "FilterContext", NetOffice.OWC10Api.PivotField.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "FilterContext", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotTotal FilterOn
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotTotal>(this, "FilterOn", NetOffice.OWC10Api.PivotTotal.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "FilterOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object FilterOnScope
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FilterOnScope");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "FilterOnScope", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object FilterFunctionValue
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FilterFunctionValue");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "FilterFunctionValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotFont GroupedFont
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFont>(this, "GroupedFont", NetOffice.OWC10Api.PivotFont.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object GroupedForeColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "GroupedForeColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "GroupedForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object GroupedBackColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "GroupedBackColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "GroupedBackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotHAlignmentEnum GroupedHAlignment
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotHAlignmentEnum>(this, "GroupedHAlignment");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "GroupedHAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotMembers CustomGroupMembers
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotMembers>(this, "CustomGroupMembers", NetOffice.OWC10Api.PivotMembers.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object DefaultValue
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultValue");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DefaultValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotFont SubtotalLabelFont
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFont>(this, "SubtotalLabelFont", NetOffice.OWC10Api.PivotFont.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object SubtotalLabelForeColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "SubtotalLabelForeColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "SubtotalLabelForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object SubtotalLabelBackColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "SubtotalLabelBackColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "SubtotalLabelBackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotHAlignmentEnum SubtotalLabelHAlignment
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotHAlignmentEnum>(this, "SubtotalLabelHAlignment");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SubtotalLabelHAlignment", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="parent">object parent</param>
		/// <param name="varChildMembers">object varChildMembers</param>
		/// <param name="bstrCaption">optional string bstrCaption = </param>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api.PivotMember AddCustomGroupMember(object parent, object varChildMembers, object bstrCaption)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api.PivotMember>(this, "AddCustomGroupMember", parent, varChildMembers, bstrCaption);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="parent">object parent</param>
		/// <param name="varChildMembers">object varChildMembers</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotMember AddCustomGroupMember(object parent, object varChildMembers)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OWC10Api.PivotMember>(this, "AddCustomGroupMember", parent, varChildMembers);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="customGroupMember">object customGroupMember</param>
		[SupportByVersion("OWC10", 1)]
		public void DeleteCustomGroupMember(object customGroupMember)
		{
			 Factory.ExecuteMethod(this, "DeleteCustomGroupMember", customGroupMember);
		}

		#endregion

		#pragma warning restore
	}
}
