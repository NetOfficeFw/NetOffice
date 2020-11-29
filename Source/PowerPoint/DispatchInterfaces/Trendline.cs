using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface Trendline 
	/// SupportByVersion PowerPoint, 14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline"/> </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Trendline : COMObject
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
                    _type = typeof(Trendline);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Trendline(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Trendline(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendline(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendline(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendline(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendline(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendline() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendline(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.Border"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.DataLabel"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.DisplayEquation"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool DisplayEquation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayEquation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayEquation", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.DisplayRSquared"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool DisplayRSquared
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayRSquared");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayRSquared", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.Index"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 Index
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.Intercept"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Double Intercept
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Intercept");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Intercept", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.InterceptIsAuto"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool InterceptIsAuto
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InterceptIsAuto");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InterceptIsAuto", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.Name"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.NameIsAuto"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public bool NameIsAuto
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "NameIsAuto");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NameIsAuto", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.Order"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 Order
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Order");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Order", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.Period"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 Period
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Period");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Period", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.Type"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Enums.XlTrendlineType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlTrendlineType>(this, "Type");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.Format"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.Backward2"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Double Backward2
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Backward2");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Backward2", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.Forward2"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Double Forward2
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Forward2");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Forward2", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.Application"/> </remarks>
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
		public Int32 Backward
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Backward");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Backward", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 Forward
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Forward");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Forward", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.ClearFormats"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object ClearFormats()
		{
			return Factory.ExecuteVariantMethodGet(this, "ClearFormats");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.Delete"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Delete()
		{
			return Factory.ExecuteVariantMethodGet(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendline.Select"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Select()
		{
			return Factory.ExecuteVariantMethodGet(this, "Select");
		}

		#endregion

		#pragma warning restore
	}
}
