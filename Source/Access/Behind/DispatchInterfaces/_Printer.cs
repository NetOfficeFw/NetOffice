using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.AccessApi;

namespace NetOffice.AccessApi.Behind
{
	/// <summary>
	/// DispatchInterface _Printer 
	/// SupportByVersion Access, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Printer : COMObject, NetOffice.AccessApi._Printer
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
                    _contractType = typeof(NetOffice.AccessApi._Printer);
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
                    _type = typeof(_Printer);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _Printer() : base()
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
		public virtual NetOffice.AccessApi.Enums.AcPrintColor ColorMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcPrintColor>(this, "ColorMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ColorMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193471.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual Int32 Copies
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Copies");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Copies", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822789.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string DeviceName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DeviceName");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195870.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string DriverName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DriverName");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198051.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Enums.AcPrintDuplex Duplex
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcPrintDuplex>(this, "Duplex");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Duplex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191910.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Enums.AcPrintOrientation Orientation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcPrintOrientation>(this, "Orientation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Orientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834798.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Enums.AcPrintPaperBin PaperBin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcPrintPaperBin>(this, "PaperBin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PaperBin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836635.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Enums.AcPrintPaperSize PaperSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcPrintPaperSize>(this, "PaperSize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PaperSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845317.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string Port
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Port");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195844.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Enums.AcPrintObjQuality PrintQuality
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcPrintObjQuality>(this, "PrintQuality");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PrintQuality", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194827.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual Int32 LeftMargin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "LeftMargin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LeftMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834469.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual Int32 RightMargin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RightMargin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RightMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835658.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual Int32 TopMargin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "TopMargin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TopMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835336.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual Int32 BottomMargin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BottomMargin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BottomMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192121.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual bool DataOnly
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DataOnly");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DataOnly", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195704.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual Int32 ItemsAcross
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ItemsAcross");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ItemsAcross", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196146.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual Int32 RowSpacing
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RowSpacing");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RowSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844923.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual Int32 ColumnSpacing
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ColumnSpacing");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ColumnSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822094.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual bool DefaultSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DefaultSize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196498.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual Int32 ItemSizeWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ItemSizeWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ItemSizeWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196765.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual Int32 ItemSizeHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ItemSizeHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ItemSizeHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194662.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Enums.AcPrintItemLayout ItemLayout
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcPrintItemLayout>(this, "ItemLayout");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ItemLayout", value);
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
		public virtual bool IsMemberSafe(Int32 dispid)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsMemberSafe", dispid);
		}

		#endregion

		#pragma warning restore
	}
}

