using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _Printer 
	/// SupportByVersion Access, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Printer : COMObject
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
                    _type = typeof(_Printer);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Printer(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Printer(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Printer(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Printer(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Printer(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Printer(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Printer() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Printer(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194552.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public NetOffice.AccessApi.Enums.AcPrintColor ColorMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcPrintColor>(this, "ColorMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ColorMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193471.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int32 Copies
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Copies");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Copies", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822789.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string DeviceName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DeviceName");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195870.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string DriverName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DriverName");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198051.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public NetOffice.AccessApi.Enums.AcPrintDuplex Duplex
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcPrintDuplex>(this, "Duplex");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Duplex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191910.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public NetOffice.AccessApi.Enums.AcPrintOrientation Orientation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcPrintOrientation>(this, "Orientation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Orientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834798.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public NetOffice.AccessApi.Enums.AcPrintPaperBin PaperBin
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcPrintPaperBin>(this, "PaperBin");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PaperBin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836635.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public NetOffice.AccessApi.Enums.AcPrintPaperSize PaperSize
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcPrintPaperSize>(this, "PaperSize");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PaperSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845317.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string Port
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Port");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195844.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public NetOffice.AccessApi.Enums.AcPrintObjQuality PrintQuality
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcPrintObjQuality>(this, "PrintQuality");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PrintQuality", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194827.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int32 LeftMargin
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LeftMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LeftMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834469.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int32 RightMargin
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RightMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RightMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835658.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int32 TopMargin
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TopMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TopMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835336.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int32 BottomMargin
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BottomMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BottomMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192121.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool DataOnly
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DataOnly");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataOnly", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195704.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int32 ItemsAcross
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ItemsAcross");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ItemsAcross", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196146.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int32 RowSpacing
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RowSpacing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RowSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844923.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int32 ColumnSpacing
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ColumnSpacing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ColumnSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822094.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool DefaultSize
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DefaultSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196498.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int32 ItemSizeWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ItemSizeWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ItemSizeWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196765.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int32 ItemSizeHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ItemSizeHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ItemSizeHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194662.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public NetOffice.AccessApi.Enums.AcPrintItemLayout ItemLayout
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcPrintItemLayout>(this, "ItemLayout");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ItemLayout", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public bool IsMemberSafe(Int32 dispid)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsMemberSafe", dispid);
		}

		#endregion

		#pragma warning restore
	}
}
