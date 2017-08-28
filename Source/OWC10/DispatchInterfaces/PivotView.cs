using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PivotView 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PivotView : COMObject
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
                    _type = typeof(PivotView);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PivotView(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PivotView(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotView(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotView(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotView(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotView(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotView() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotView(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotFieldSets FieldSets
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFieldSets>(this, "FieldSets", NetOffice.OWC10Api.PivotFieldSets.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotGroupAxis RowAxis
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotGroupAxis>(this, "RowAxis", NetOffice.OWC10Api.PivotGroupAxis.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotGroupAxis ColumnAxis
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotGroupAxis>(this, "ColumnAxis", NetOffice.OWC10Api.PivotGroupAxis.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotDataAxis DataAxis
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotDataAxis>(this, "DataAxis", NetOffice.OWC10Api.PivotDataAxis.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotFilterAxis FilterAxis
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFilterAxis>(this, "FilterAxis", NetOffice.OWC10Api.PivotFilterAxis.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotLabel Label
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotLabel>(this, "Label", NetOffice.OWC10Api.PivotLabel.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotLabel TitleBar
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotLabel>(this, "TitleBar", NetOffice.OWC10Api.PivotLabel.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotTotals Totals
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotTotals>(this, "Totals", NetOffice.OWC10Api.PivotTotals.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotFont TotalFont
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFont>(this, "TotalFont", NetOffice.OWC10Api.PivotFont.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object TotalForeColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TotalForeColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "TotalForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object TotalBackColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TotalBackColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "TotalBackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotFont HeaderFont
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFont>(this, "HeaderFont", NetOffice.OWC10Api.PivotFont.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object HeaderForeColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "HeaderForeColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "HeaderForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object HeaderBackColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "HeaderBackColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "HeaderBackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotHAlignmentEnum HeaderHAlignment
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotHAlignmentEnum>(this, "HeaderHAlignment");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "HeaderHAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 HeaderHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HeaderHeight");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotFont FieldLabelFont
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFont>(this, "FieldLabelFont", NetOffice.OWC10Api.PivotFont.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object FieldLabelForeColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FieldLabelForeColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "FieldLabelForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object FieldLabelBackColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FieldLabelBackColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "FieldLabelBackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 FieldLabelHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "FieldLabelHeight");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 DetailRowHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DetailRowHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DetailRowHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public object DetailSortOrder
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DetailSortOrder");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DetailSortOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotViewTotalOrientationEnum TotalOrientation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotViewTotalOrientationEnum>(this, "TotalOrientation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TotalOrientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool TotalAllMembers
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TotalAllMembers");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TotalAllMembers", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 DetailMaxWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DetailMaxWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DetailMaxWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 DetailMaxHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DetailMaxHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DetailMaxHeight", value);
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
		public bool IsFiltered
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsFiltered");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IsFiltered", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayCalculatedMembers
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayCalculatedMembers");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayCalculatedMembers", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool UseProviderFormatting
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseProviderFormatting");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseProviderFormatting", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotTableExpandEnum ExpandDetails
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotTableExpandEnum>(this, "ExpandDetails");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ExpandDetails", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api.IPivotControl Control
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.IPivotControl>(this, "Control");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotGroupAxis PageAxis
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotGroupAxis>(this, "PageAxis", NetOffice.OWC10Api.PivotGroupAxis.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotTableExpandEnum ExpandMembers
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotTableExpandEnum>(this, "ExpandMembers");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ExpandMembers", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowEdits
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowEdits");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowEdits", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowAdditions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowAdditions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowAdditions", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowDeletions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowDeletions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowDeletions", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotFont PropertyCaptionFont
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFont>(this, "PropertyCaptionFont", NetOffice.OWC10Api.PivotFont.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotFont PropertyValueFont
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFont>(this, "PropertyValueFont", NetOffice.OWC10Api.PivotFont.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotHAlignmentEnum PropertyCaptionHAlignment
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotHAlignmentEnum>(this, "PropertyCaptionHAlignment");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PropertyCaptionHAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.PivotHAlignmentEnum PropertyValueHAlignment
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotHAlignmentEnum>(this, "PropertyValueHAlignment");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PropertyValueHAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool DisplayCellColor
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayCellColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayCellColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool FilterCrossJoins
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FilterCrossJoins");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FilterCrossJoins", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="field">NetOffice.OWC10Api.PivotField field</param>
		/// <param name="function">NetOffice.OWC10Api.Enums.PivotTotalFunctionEnum function</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotTotal AddTotal(string name, NetOffice.OWC10Api.PivotField field, NetOffice.OWC10Api.Enums.PivotTotalFunctionEnum function)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PivotTotal>(this, "AddTotal", NetOffice.OWC10Api.PivotTotal.LateBindingApiWrapperType, name, field, function);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="total">object total</param>
		[SupportByVersion("OWC10", 1)]
		public void DeleteTotal(object total)
		{
			 Factory.ExecuteMethod(this, "DeleteTotal", total);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotFieldSet AddFieldSet(string name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PivotFieldSet>(this, "AddFieldSet", NetOffice.OWC10Api.PivotFieldSet.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="fieldSet">object fieldSet</param>
		[SupportByVersion("OWC10", 1)]
		public void DeleteFieldSet(object fieldSet)
		{
			 Factory.ExecuteMethod(this, "DeleteFieldSet", fieldSet);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="maxDataFields">optional Int32 MaxDataFields = 0</param>
		[SupportByVersion("OWC10", 1)]
		public void AutoLayout(object maxDataFields)
		{
			 Factory.ExecuteMethod(this, "AutoLayout", maxDataFields);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void AutoLayout()
		{
			 Factory.ExecuteMethod(this, "AutoLayout");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="caption">string caption</param>
		/// <param name="expression">string expression</param>
		/// <param name="solveOrder">optional Int32 SolveOrder = 0</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotTotal AddCalculatedTotal(string name, string caption, string expression, object solveOrder)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PivotTotal>(this, "AddCalculatedTotal", NetOffice.OWC10Api.PivotTotal.LateBindingApiWrapperType, name, caption, expression, solveOrder);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="caption">string caption</param>
		/// <param name="expression">string expression</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PivotTotal AddCalculatedTotal(string name, string caption, string expression)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PivotTotal>(this, "AddCalculatedTotal", NetOffice.OWC10Api.PivotTotal.LateBindingApiWrapperType, name, caption, expression);
		}

		#endregion

		#pragma warning restore
	}
}
