using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Axis 
	/// SupportByVersion Word, 14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis"/> </remarks>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Axis : COMObject
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
                    _type = typeof(Axis);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Axis(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Axis(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Axis(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Axis(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Axis(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Axis(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Axis() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Axis(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.AxisBetweenCategories"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool AxisBetweenCategories
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AxisBetweenCategories");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AxisBetweenCategories", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.AxisGroup"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlAxisGroup AxisGroup
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlAxisGroup>(this, "AxisGroup");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.AxisTitle"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.AxisTitle AxisTitle
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.AxisTitle>(this, "AxisTitle", NetOffice.WordApi.AxisTitle.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.CategoryNames"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public object CategoryNames
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CategoryNames");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "CategoryNames", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.Crosses"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlAxisCrosses Crosses
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlAxisCrosses>(this, "Crosses");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Crosses", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.CrossesAt"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Double CrossesAt
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "CrossesAt");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CrossesAt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.HasMajorGridlines"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool HasMajorGridlines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasMajorGridlines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasMajorGridlines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.HasMinorGridlines"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool HasMinorGridlines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasMinorGridlines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasMinorGridlines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.HasTitle"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool HasTitle
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasTitle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasTitle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.MajorGridlines"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Gridlines MajorGridlines
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Gridlines>(this, "MajorGridlines", NetOffice.WordApi.Gridlines.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.MajorTickMark"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlTickMark MajorTickMark
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlTickMark>(this, "MajorTickMark");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MajorTickMark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.MajorUnit"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Double MajorUnit
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "MajorUnit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MajorUnit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.LogBase"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Double LogBase
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "LogBase");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LogBase", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.TickLabelSpacingIsAuto"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool TickLabelSpacingIsAuto
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TickLabelSpacingIsAuto");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TickLabelSpacingIsAuto", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.MajorUnitIsAuto"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool MajorUnitIsAuto
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MajorUnitIsAuto");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MajorUnitIsAuto", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.MaximumScale"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Double MaximumScale
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "MaximumScale");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaximumScale", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.MaximumScaleIsAuto"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool MaximumScaleIsAuto
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MaximumScaleIsAuto");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaximumScaleIsAuto", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.MinimumScale"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Double MinimumScale
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "MinimumScale");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MinimumScale", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.MinimumScaleIsAuto"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool MinimumScaleIsAuto
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MinimumScaleIsAuto");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MinimumScaleIsAuto", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.MinorGridlines"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Gridlines MinorGridlines
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Gridlines>(this, "MinorGridlines", NetOffice.WordApi.Gridlines.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.MinorTickMark"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlTickMark MinorTickMark
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlTickMark>(this, "MinorTickMark");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MinorTickMark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.MinorUnit"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Double MinorUnit
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "MinorUnit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MinorUnit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.MinorUnitIsAuto"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool MinorUnitIsAuto
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MinorUnitIsAuto");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MinorUnitIsAuto", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.ReversePlotOrder"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool ReversePlotOrder
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReversePlotOrder");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReversePlotOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.ScaleType"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlScaleType ScaleType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlScaleType>(this, "ScaleType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ScaleType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.TickLabelPosition"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlTickLabelPosition TickLabelPosition
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlTickLabelPosition>(this, "TickLabelPosition");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TickLabelPosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.TickLabels"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.TickLabels TickLabels
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TickLabels>(this, "TickLabels", NetOffice.WordApi.TickLabels.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.TickLabelSpacing"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Int32 TickLabelSpacing
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TickLabelSpacing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TickLabelSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.TickMarkSpacing"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Int32 TickMarkSpacing
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TickMarkSpacing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TickMarkSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.Type"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlAxisType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlAxisType>(this, "Type");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.BaseUnit"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlTimeUnit BaseUnit
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlTimeUnit>(this, "BaseUnit");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BaseUnit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.BaseUnitIsAuto"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool BaseUnitIsAuto
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BaseUnitIsAuto");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BaseUnitIsAuto", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.MajorUnitScale"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlTimeUnit MajorUnitScale
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlTimeUnit>(this, "MajorUnitScale");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MajorUnitScale", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.MinorUnitScale"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlTimeUnit MinorUnitScale
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlTimeUnit>(this, "MinorUnitScale");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MinorUnitScale", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.CategoryType"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlCategoryType CategoryType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlCategoryType>(this, "CategoryType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CategoryType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.Left"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Double Left
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Left");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.Top"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Double Top
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Top");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.Width"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Double Width
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Width");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.Height"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Double Height
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Height");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.DisplayUnit"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.XlDisplayUnit DisplayUnit
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlDisplayUnit>(this, "DisplayUnit");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DisplayUnit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.DisplayUnitCustom"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Double DisplayUnitCustom
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "DisplayUnitCustom");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayUnitCustom", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.HasDisplayUnitLabel"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool HasDisplayUnitLabel
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasDisplayUnitLabel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasDisplayUnitLabel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.DisplayUnitLabel"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.DisplayUnitLabel DisplayUnitLabel
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.DisplayUnitLabel>(this, "DisplayUnitLabel", NetOffice.WordApi.DisplayUnitLabel.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.Border"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ChartBorder Border
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartBorder>(this, "Border", NetOffice.WordApi.ChartBorder.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.Format"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ChartFormat Format
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartFormat>(this, "Format", NetOffice.WordApi.ChartFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.Application"/> </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		public object Application
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.Creator"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.Parent"/> </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.Delete"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public object Delete()
		{
			return Factory.ExecuteVariantMethodGet(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Axis.Select"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public object Select()
		{
			return Factory.ExecuteVariantMethodGet(this, "Select");
		}

		#endregion

		#pragma warning restore
	}
}
