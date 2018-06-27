using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface PivotView 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PivotView : COMObject, NetOffice.OWC10Api.PivotView
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
                    _contractType = typeof(NetOffice.OWC10Api.PivotView);
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
                    _type = typeof(PivotView);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public PivotView() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotFieldSets FieldSets
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFieldSets>(this, "FieldSets", typeof(NetOffice.OWC10Api.PivotFieldSets));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotGroupAxis RowAxis
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotGroupAxis>(this, "RowAxis", typeof(NetOffice.OWC10Api.PivotGroupAxis));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotGroupAxis ColumnAxis
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotGroupAxis>(this, "ColumnAxis", typeof(NetOffice.OWC10Api.PivotGroupAxis));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotDataAxis DataAxis
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotDataAxis>(this, "DataAxis", typeof(NetOffice.OWC10Api.PivotDataAxis));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotFilterAxis FilterAxis
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFilterAxis>(this, "FilterAxis", typeof(NetOffice.OWC10Api.PivotFilterAxis));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotLabel Label
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotLabel>(this, "Label", typeof(NetOffice.OWC10Api.PivotLabel));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotLabel TitleBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotLabel>(this, "TitleBar", typeof(NetOffice.OWC10Api.PivotLabel));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotTotals Totals
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotTotals>(this, "Totals", typeof(NetOffice.OWC10Api.PivotTotals));			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotFont TotalFont
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFont>(this, "TotalFont", typeof(NetOffice.OWC10Api.PivotFont));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual object TotalForeColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "TotalForeColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "TotalForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual object TotalBackColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "TotalBackColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "TotalBackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotFont HeaderFont
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFont>(this, "HeaderFont", typeof(NetOffice.OWC10Api.PivotFont));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual object HeaderForeColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "HeaderForeColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "HeaderForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual object HeaderBackColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "HeaderBackColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "HeaderBackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.PivotHAlignmentEnum HeaderHAlignment
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotHAlignmentEnum>(this, "HeaderHAlignment");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "HeaderHAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 HeaderHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "HeaderHeight");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotFont FieldLabelFont
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFont>(this, "FieldLabelFont", typeof(NetOffice.OWC10Api.PivotFont));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual object FieldLabelForeColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FieldLabelForeColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "FieldLabelForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual object FieldLabelBackColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FieldLabelBackColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "FieldLabelBackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 FieldLabelHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "FieldLabelHeight");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 DetailRowHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DetailRowHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DetailRowHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual object DetailSortOrder
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DetailSortOrder");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "DetailSortOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.PivotViewTotalOrientationEnum TotalOrientation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotViewTotalOrientationEnum>(this, "TotalOrientation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "TotalOrientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool TotalAllMembers
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "TotalAllMembers");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TotalAllMembers", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 DetailMaxWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DetailMaxWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DetailMaxWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 DetailMaxHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DetailMaxHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DetailMaxHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool DetailAutoFit
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DetailAutoFit");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DetailAutoFit", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool IsFiltered
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsFiltered");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IsFiltered", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool DisplayCalculatedMembers
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayCalculatedMembers");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayCalculatedMembers", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool UseProviderFormatting
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UseProviderFormatting");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UseProviderFormatting", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.PivotTableExpandEnum ExpandDetails
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotTableExpandEnum>(this, "ExpandDetails");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ExpandDetails", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public virtual NetOffice.OWC10Api.IPivotControl Control
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.IPivotControl>(this, "Control");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotGroupAxis PageAxis
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotGroupAxis>(this, "PageAxis", typeof(NetOffice.OWC10Api.PivotGroupAxis));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.PivotTableExpandEnum ExpandMembers
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotTableExpandEnum>(this, "ExpandMembers");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ExpandMembers", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AllowEdits
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowEdits");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowEdits", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AllowAdditions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowAdditions");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowAdditions", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AllowDeletions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowDeletions");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowDeletions", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotFont PropertyCaptionFont
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFont>(this, "PropertyCaptionFont", typeof(NetOffice.OWC10Api.PivotFont));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotFont PropertyValueFont
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFont>(this, "PropertyValueFont", typeof(NetOffice.OWC10Api.PivotFont));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.PivotHAlignmentEnum PropertyCaptionHAlignment
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotHAlignmentEnum>(this, "PropertyCaptionHAlignment");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PropertyCaptionHAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.PivotHAlignmentEnum PropertyValueHAlignment
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotHAlignmentEnum>(this, "PropertyValueHAlignment");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PropertyValueHAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool DisplayCellColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayCellColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayCellColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool FilterCrossJoins
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FilterCrossJoins");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FilterCrossJoins", value);
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
		public virtual NetOffice.OWC10Api.PivotTotal AddTotal(string name, NetOffice.OWC10Api.PivotField field, NetOffice.OWC10Api.Enums.PivotTotalFunctionEnum function)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PivotTotal>(this, "AddTotal", typeof(NetOffice.OWC10Api.PivotTotal), name, field, function);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="total">object total</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void DeleteTotal(object total)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteTotal", total);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotFieldSet AddFieldSet(string name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PivotFieldSet>(this, "AddFieldSet", typeof(NetOffice.OWC10Api.PivotFieldSet), name);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="fieldSet">object fieldSet</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void DeleteFieldSet(object fieldSet)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteFieldSet", fieldSet);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="maxDataFields">optional Int32 MaxDataFields = 0</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void AutoLayout(object maxDataFields)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AutoLayout", maxDataFields);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void AutoLayout()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AutoLayout");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="caption">string caption</param>
		/// <param name="expression">string expression</param>
		/// <param name="solveOrder">optional Int32 SolveOrder = 0</param>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotTotal AddCalculatedTotal(string name, string caption, string expression, object solveOrder)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PivotTotal>(this, "AddCalculatedTotal", typeof(NetOffice.OWC10Api.PivotTotal), name, caption, expression, solveOrder);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="caption">string caption</param>
		/// <param name="expression">string expression</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotTotal AddCalculatedTotal(string name, string caption, string expression)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PivotTotal>(this, "AddCalculatedTotal", typeof(NetOffice.OWC10Api.PivotTotal), name, caption, expression);
		}

		#endregion

		#pragma warning restore
	}
}


