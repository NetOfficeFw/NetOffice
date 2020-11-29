using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface Point 
	/// SupportByVersion PowerPoint, 14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point"/> </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Point : COMObject
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
                    _type = typeof(Point);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Point(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Point(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Point(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Point(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Point(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Point(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Point() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Point(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.Parent"/> </remarks>
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
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.DataLabel"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.DataLabel DataLabel
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.DataLabel>(this, "DataLabel", NetOffice.PowerPointApi.DataLabel.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.Explosion"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 Explosion
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Explosion");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Explosion", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.HasDataLabel"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool HasDataLabel
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasDataLabel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasDataLabel", value);
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
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Interior>(this, "Interior", NetOffice.PowerPointApi.Interior.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.InvertIfNegative"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool InvertIfNegative
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InvertIfNegative");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InvertIfNegative", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.MarkerBackgroundColor"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 MarkerBackgroundColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MarkerBackgroundColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarkerBackgroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.MarkerBackgroundColorIndex"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlColorIndex MarkerBackgroundColorIndex
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlColorIndex>(this, "MarkerBackgroundColorIndex");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MarkerBackgroundColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.MarkerForegroundColor"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 MarkerForegroundColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MarkerForegroundColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarkerForegroundColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.MarkerForegroundColorIndex"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlColorIndex MarkerForegroundColorIndex
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlColorIndex>(this, "MarkerForegroundColorIndex");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MarkerForegroundColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.MarkerSize"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 MarkerSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MarkerSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarkerSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.MarkerStyle"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlMarkerStyle MarkerStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlMarkerStyle>(this, "MarkerStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MarkerStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.PictureType"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlChartPictureType PictureType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlChartPictureType>(this, "PictureType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PictureType", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.ApplyPictToSides"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool ApplyPictToSides
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ApplyPictToSides");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ApplyPictToSides", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.ApplyPictToFront"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool ApplyPictToFront
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ApplyPictToFront");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ApplyPictToFront", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.ApplyPictToEnd"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool ApplyPictToEnd
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ApplyPictToEnd");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ApplyPictToEnd", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.Shadow"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool Shadow
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Shadow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Shadow", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.SecondaryPlot"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool SecondaryPlot
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SecondaryPlot");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SecondaryPlot", value);
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
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartFillFormat>(this, "Fill", NetOffice.PowerPointApi.ChartFillFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.Has3DEffect"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool Has3DEffect
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Has3DEffect");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Has3DEffect", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.Format"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.Creator"/> </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.PictureUnit2"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Double PictureUnit2
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "PictureUnit2");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureUnit2", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.Application"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", NetOffice.PowerPointApi.Application.LateBindingApiWrapperType);
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
				return Factory.ExecuteInt32PropertyGet(this, "PictureUnit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureUnit", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.Name"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.Height"/> </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.Width"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.Left"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.Top"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Double Top
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Top");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.ClearFormats"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ClearFormats()
		{
			return Factory.ExecuteVariantMethodGet(this, "ClearFormats");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.Copy"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Copy()
		{
			return Factory.ExecuteVariantMethodGet(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.Delete"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Delete()
		{
			return Factory.ExecuteVariantMethodGet(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.Paste"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Paste()
		{
			return Factory.ExecuteVariantMethodGet(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.Select"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Select()
		{
			return Factory.ExecuteVariantMethodGet(this, "Select");
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
			return Factory.ExecuteVariantMethodGet(this, "_ApplyDataLabels", type, legendKey, autoText, hasLeaderLines);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object _ApplyDataLabels()
		{
			return Factory.ExecuteVariantMethodGet(this, "_ApplyDataLabels");
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
			return Factory.ExecuteVariantMethodGet(this, "_ApplyDataLabels", type);
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
			return Factory.ExecuteVariantMethodGet(this, "_ApplyDataLabels", type, legendKey);
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
			return Factory.ExecuteVariantMethodGet(this, "_ApplyDataLabels", type, legendKey, autoText);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.ApplyDataLabels"/> </remarks>
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
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize, separator });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.ApplyDataLabels"/> </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ApplyDataLabels()
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.ApplyDataLabels"/> </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ApplyDataLabels(object type)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.ApplyDataLabels"/> </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", type, legendKey);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.ApplyDataLabels"/> </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", type, legendKey, autoText);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.ApplyDataLabels"/> </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", type, legendKey, autoText, hasLeaderLines);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.ApplyDataLabels"/> </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
		/// <param name="legendKey">optional object legendKey</param>
		/// <param name="autoText">optional object autoText</param>
		/// <param name="hasLeaderLines">optional object hasLeaderLines</param>
		/// <param name="showSeriesName">optional object showSeriesName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName)
		{
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.ApplyDataLabels"/> </remarks>
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
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.ApplyDataLabels"/> </remarks>
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
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.ApplyDataLabels"/> </remarks>
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
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.ApplyDataLabels"/> </remarks>
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
			return Factory.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[]{ type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.PieSliceLocation"/> </remarks>
		/// <param name="loc">NetOffice.PowerPointApi.Enums.XlPieSliceLocation loc</param>
		/// <param name="index">optional NetOffice.PowerPointApi.Enums.XlPieSliceIndex Index = 2</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Double PieSliceLocation(NetOffice.PowerPointApi.Enums.XlPieSliceLocation loc, object index)
		{
			return Factory.ExecuteDoubleMethodGet(this, "PieSliceLocation", loc, index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Point.PieSliceLocation"/> </remarks>
		/// <param name="loc">NetOffice.PowerPointApi.Enums.XlPieSliceLocation loc</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Double PieSliceLocation(NetOffice.PowerPointApi.Enums.XlPieSliceLocation loc)
		{
			return Factory.ExecuteDoubleMethodGet(this, "PieSliceLocation", loc);
		}

		#endregion

		#pragma warning restore
	}
}
