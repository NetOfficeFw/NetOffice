using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface Axis 
	/// SupportByVersion PowerPoint, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744572.aspx </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745589.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744123.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlAxisGroup AxisGroup
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlAxisGroup>(this, "AxisGroup");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746159.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.AxisTitle AxisTitle
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.AxisTitle>(this, "AxisTitle", NetOffice.PowerPointApi.AxisTitle.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746687.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745651.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlAxisCrosses Crosses
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlAxisCrosses>(this, "Crosses");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Crosses", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746282.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745884.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744847.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743898.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746330.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Gridlines MajorGridlines
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Gridlines>(this, "MajorGridlines", NetOffice.PowerPointApi.Gridlines.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745421.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlTickMark MajorTickMark
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlTickMark>(this, "MajorTickMark");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MajorTickMark", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745026.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746487.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746691.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746838.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746268.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746704.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745624.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745391.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746773.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Gridlines MinorGridlines
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Gridlines>(this, "MinorGridlines", NetOffice.PowerPointApi.Gridlines.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744253.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlTickMark MinorTickMark
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlTickMark>(this, "MinorTickMark");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MinorTickMark", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746830.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744117.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745065.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746084.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlScaleType ScaleType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlScaleType>(this, "ScaleType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ScaleType", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744718.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlTickLabelPosition TickLabelPosition
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlTickLabelPosition>(this, "TickLabelPosition");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TickLabelPosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745407.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.TickLabels TickLabels
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.TickLabels>(this, "TickLabels", NetOffice.PowerPointApi.TickLabels.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745736.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745472.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746690.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlAxisType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlAxisType>(this, "Type");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745843.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlTimeUnit BaseUnit
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlTimeUnit>(this, "BaseUnit");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BaseUnit", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744623.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744714.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlTimeUnit MajorUnitScale
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlTimeUnit>(this, "MajorUnitScale");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MajorUnitScale", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744083.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlTimeUnit MinorUnitScale
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlTimeUnit>(this, "MinorUnitScale");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MinorUnitScale", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746102.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlCategoryType CategoryType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlCategoryType>(this, "CategoryType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CategoryType", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746324.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Double Left
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Left");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743906.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Double Top
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Top");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745897.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Double Width
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Width");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746380.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Double Height
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Height");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745098.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlDisplayUnit DisplayUnit
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlDisplayUnit>(this, "DisplayUnit");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DisplayUnit", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746148.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745935.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745287.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.DisplayUnitLabel DisplayUnitLabel
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.DisplayUnitLabel>(this, "DisplayUnitLabel", NetOffice.PowerPointApi.DisplayUnitLabel.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746825.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.ChartBorder Border
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartBorder>(this, "Border", NetOffice.PowerPointApi.ChartBorder.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746150.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.ChartFormat Format
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartFormat>(this, "Format", NetOffice.PowerPointApi.ChartFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745944.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744067.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746674.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", NetOffice.PowerPointApi.Application.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746130.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Delete()
		{
			return Factory.ExecuteVariantMethodGet(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744150.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Select()
		{
			return Factory.ExecuteVariantMethodGet(this, "Select");
		}

		#endregion

		#pragma warning restore
	}
}
