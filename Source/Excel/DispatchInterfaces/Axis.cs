using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface Axis 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis(object)"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.AxisBetweenCategories"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.AxisGroup"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlAxisGroup AxisGroup
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlAxisGroup>(this, "AxisGroup");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.AxisTitle"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.AxisTitle AxisTitle
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.AxisTitle>(this, "AxisTitle", NetOffice.ExcelApi.AxisTitle.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.Border"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Border Border
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Border>(this, "Border", NetOffice.ExcelApi.Border.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.CategoryNames"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.Crosses"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlAxisCrosses Crosses
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlAxisCrosses>(this, "Crosses");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Crosses", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.CrossesAt"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.HasMajorGridlines"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.HasMinorGridlines"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.HasTitle"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.MajorGridlines"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Gridlines MajorGridlines
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Gridlines>(this, "MajorGridlines", NetOffice.ExcelApi.Gridlines.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.MajorTickMark"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlTickMark MajorTickMark
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTickMark>(this, "MajorTickMark");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MajorTickMark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.MajorUnit"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.MajorUnitIsAuto"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.MaximumScale"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.MaximumScaleIsAuto"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.MinimumScale"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.MinimumScaleIsAuto"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.MinorGridlines"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Gridlines MinorGridlines
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Gridlines>(this, "MinorGridlines", NetOffice.ExcelApi.Gridlines.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.MinorTickMark"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlTickMark MinorTickMark
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTickMark>(this, "MinorTickMark");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MinorTickMark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.MinorUnit"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.MinorUnitIsAuto"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.ReversePlotOrder"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.ScaleType"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlScaleType ScaleType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlScaleType>(this, "ScaleType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ScaleType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.TickLabelPosition"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlTickLabelPosition TickLabelPosition
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTickLabelPosition>(this, "TickLabelPosition");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TickLabelPosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.TickLabels"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.TickLabels TickLabels
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.TickLabels>(this, "TickLabels", NetOffice.ExcelApi.TickLabels.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.TickLabelSpacing"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.TickMarkSpacing"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.Type"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlAxisType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlAxisType>(this, "Type");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.BaseUnit"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlTimeUnit BaseUnit
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTimeUnit>(this, "BaseUnit");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BaseUnit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.BaseUnitIsAuto"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.MajorUnitScale"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlTimeUnit MajorUnitScale
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTimeUnit>(this, "MajorUnitScale");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MajorUnitScale", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.MinorUnitScale"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlTimeUnit MinorUnitScale
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTimeUnit>(this, "MinorUnitScale");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MinorUnitScale", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.CategoryType"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCategoryType CategoryType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCategoryType>(this, "CategoryType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CategoryType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.Left"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Double Left
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Left");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.Top"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Double Top
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Top");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.Width"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Double Width
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Width");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.Height"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Double Height
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Height");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.DisplayUnit"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlDisplayUnit DisplayUnit
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlDisplayUnit>(this, "DisplayUnit");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DisplayUnit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.DisplayUnitCustom"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.HasDisplayUnitLabel"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.DisplayUnitLabel"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.DisplayUnitLabel DisplayUnitLabel
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.DisplayUnitLabel>(this, "DisplayUnitLabel", NetOffice.ExcelApi.DisplayUnitLabel.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.LogBase"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.TickLabelSpacingIsAuto"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.Format"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.ChartFormat Format
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ChartFormat>(this, "Format", NetOffice.ExcelApi.ChartFormat.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.Delete"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Delete()
		{
			return Factory.ExecuteVariantMethodGet(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Axis.Select"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Select()
		{
			return Factory.ExecuteVariantMethodGet(this, "Select");
		}

		#endregion

		#pragma warning restore
	}
}
