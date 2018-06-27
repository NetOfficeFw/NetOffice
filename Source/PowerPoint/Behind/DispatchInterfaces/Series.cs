using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PowerPointApi;

namespace NetOffice.PowerPointApi.Behind
{
	/// <summary>
	/// DispatchInterface Series 
	/// SupportByVersion PowerPoint, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744981.aspx </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Series : COMObject, NetOffice.PowerPointApi.Series
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
                    _contractType = typeof(NetOffice.PowerPointApi.Series);
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
                    _type = typeof(Series);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Series() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744076.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746153.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlAxisGroup AxisGroup
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlAxisGroup>(this, "AxisGroup");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "AxisGroup", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PowerPointApi.ChartBorder Border
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartBorder>(this, "Border", typeof(NetOffice.PowerPointApi.ChartBorder));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745197.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.ErrorBars ErrorBars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ErrorBars>(this, "ErrorBars", typeof(NetOffice.PowerPointApi.ErrorBars));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746257.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 Explosion
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Explosion");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Explosion", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743895.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public string Formula
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Formula");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Formula", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745659.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public string FormulaLocal
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FormulaLocal");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FormulaLocal", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744271.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public string FormulaR1C1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FormulaR1C1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FormulaR1C1", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746267.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public string FormulaR1C1Local
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FormulaR1C1Local");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FormulaR1C1Local", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745965.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool HasDataLabels
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasDataLabels");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasDataLabels", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745102.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool HasErrorBars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasErrorBars");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasErrorBars", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PowerPointApi.Interior Interior
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Interior>(this, "Interior", typeof(NetOffice.PowerPointApi.Interior));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PowerPointApi.ChartFillFormat Fill
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartFillFormat>(this, "Fill", typeof(NetOffice.PowerPointApi.ChartFillFormat));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746458.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool InvertIfNegative
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "InvertIfNegative");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InvertIfNegative", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745192.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 MarkerBackgroundColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "MarkerBackgroundColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MarkerBackgroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744114.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlColorIndex MarkerBackgroundColorIndex
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlColorIndex>(this, "MarkerBackgroundColorIndex");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "MarkerBackgroundColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744625.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 MarkerForegroundColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "MarkerForegroundColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MarkerForegroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745462.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlColorIndex MarkerForegroundColorIndex
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlColorIndex>(this, "MarkerForegroundColorIndex");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "MarkerForegroundColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745041.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 MarkerSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "MarkerSize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MarkerSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746599.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlMarkerStyle MarkerStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlMarkerStyle>(this, "MarkerStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "MarkerStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745449.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744028.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlChartPictureType PictureType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlChartPictureType>(this, "PictureType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PictureType", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744124.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 PlotOrder
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PlotOrder");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PlotOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746839.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool Smooth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Smooth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Smooth", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745492.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 Type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Type");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744368.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.OfficeApi.Enums.XlChartType ChartType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlChartType>(this, "ChartType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ChartType", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746833.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Values
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Values");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Values", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746506.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object XValues
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "XValues");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "XValues", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746198.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object BubbleSizes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BubbleSizes");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BubbleSizes", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746220.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlBarShape BarShape
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlBarShape>(this, "BarShape");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "BarShape", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746046.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool ApplyPictToSides
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ApplyPictToSides");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ApplyPictToSides", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746078.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool ApplyPictToFront
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ApplyPictToFront");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ApplyPictToFront", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746783.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool ApplyPictToEnd
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ApplyPictToEnd");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ApplyPictToEnd", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746298.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool Has3DEffect
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Has3DEffect");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Has3DEffect", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744807.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool Shadow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Shadow");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Shadow", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744802.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool HasLeaderLines
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasLeaderLines");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasLeaderLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746727.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.LeaderLines LeaderLines
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.LeaderLines>(this, "LeaderLines", typeof(NetOffice.PowerPointApi.LeaderLines));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744341.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.ChartFormat Format
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartFormat>(this, "Format", typeof(NetOffice.PowerPointApi.ChartFormat));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744543.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745441.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Double PictureUnit2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "PictureUnit2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureUnit2", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746564.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", typeof(NetOffice.PowerPointApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 PictureUnit
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PictureUnit");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureUnit", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745452.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 PlotColorIndex
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PlotColorIndex");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746517.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 InvertColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "InvertColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InvertColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745489.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlColorIndex InvertColorIndex
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlColorIndex>(this, "InvertColorIndex");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "InvertColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227564.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		public bool IsFiltered
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743916.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ClearFormats()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ClearFormats");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745312.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Copy()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746505.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object DataLabels(object index)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DataLabels", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746505.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object DataLabels()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DataLabels");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744542.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Delete()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745822.aspx </remarks>
		/// <param name="direction">NetOffice.PowerPointApi.Enums.XlErrorBarDirection direction</param>
		/// <param name="include">NetOffice.PowerPointApi.Enums.XlErrorBarInclude include</param>
		/// <param name="type">NetOffice.PowerPointApi.Enums.XlErrorBarType type</param>
		/// <param name="amount">optional object amount</param>
		/// <param name="minusValues">optional object minusValues</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ErrorBar(NetOffice.PowerPointApi.Enums.XlErrorBarDirection direction, NetOffice.PowerPointApi.Enums.XlErrorBarInclude include, NetOffice.PowerPointApi.Enums.XlErrorBarType type, object amount, object minusValues)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ErrorBar", new object[]{ direction, include, type, amount, minusValues });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745822.aspx </remarks>
		/// <param name="direction">NetOffice.PowerPointApi.Enums.XlErrorBarDirection direction</param>
		/// <param name="include">NetOffice.PowerPointApi.Enums.XlErrorBarInclude include</param>
		/// <param name="type">NetOffice.PowerPointApi.Enums.XlErrorBarType type</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ErrorBar(NetOffice.PowerPointApi.Enums.XlErrorBarDirection direction, NetOffice.PowerPointApi.Enums.XlErrorBarInclude include, NetOffice.PowerPointApi.Enums.XlErrorBarType type)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ErrorBar", direction, include, type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745822.aspx </remarks>
		/// <param name="direction">NetOffice.PowerPointApi.Enums.XlErrorBarDirection direction</param>
		/// <param name="include">NetOffice.PowerPointApi.Enums.XlErrorBarInclude include</param>
		/// <param name="type">NetOffice.PowerPointApi.Enums.XlErrorBarType type</param>
		/// <param name="amount">optional object amount</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ErrorBar(NetOffice.PowerPointApi.Enums.XlErrorBarDirection direction, NetOffice.PowerPointApi.Enums.XlErrorBarInclude include, NetOffice.PowerPointApi.Enums.XlErrorBarType type, object amount)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ErrorBar", direction, include, type, amount);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744657.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Paste()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744898.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Points(object index)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Points", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744898.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Points()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Points");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744062.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Select()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Select");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744102.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Trendlines(object index)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Trendlines", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744102.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Trendlines()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Trendlines");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyCustomType", chartType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object _ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_ApplyDataLabels", type, legendKey, autoText, hasLeaderLines);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object _ApplyDataLabels()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_ApplyDataLabels");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object _ApplyDataLabels(object type)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_ApplyDataLabels", type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object _ApplyDataLabels(object type, object legendKey)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_ApplyDataLabels", type, legendKey);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object _ApplyDataLabels(object type, object legendKey, object autoText)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_ApplyDataLabels", type, legendKey, autoText);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746413.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		/// <param name="showValue">optional object showValue</param>
		/// <param name="showPercentage">optional object showPercentage</param>
		/// <param name="showBubbleSize">optional object showBubbleSize</param>
		/// <param name="separator">optional object separator</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize, object separator)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize, separator });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746413.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ApplyDataLabels()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746413.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ApplyDataLabels(object type)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746413.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", type, legendKey);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746413.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", type, legendKey, autoText);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746413.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", type, legendKey, autoText, hasLeaderLines);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746413.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746413.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746413.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		/// <param name="showValue">optional object showValue</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746413.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		/// <param name="showValue">optional object showValue</param>
		/// <param name="showPercentage">optional object showPercentage</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746413.aspx </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		/// <param name="showCategoryName">optional object showCategoryName</param>
		/// <param name="showValue">optional object showValue</param>
		/// <param name="showPercentage">optional object showPercentage</param>
		/// <param name="showBubbleSize">optional object showBubbleSize</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize });
		}

		#endregion

		#pragma warning restore
	}
}


