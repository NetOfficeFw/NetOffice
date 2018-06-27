using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSComctlLibApi;

namespace NetOffice.MSComctlLibApi.Behind
{
	/// <summary>
	/// DispatchInterface ISlider 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class ISlider : COMObject, NetOffice.MSComctlLibApi.ISlider
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
                    _contractType = typeof(NetOffice.MSComctlLibApi.ISlider);
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
                    _type = typeof(ISlider);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ISlider() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual Int32 _Value
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "_Value");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "_Value", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual Int32 LargeChange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "LargeChange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LargeChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual Int32 SmallChange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SmallChange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SmallChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual Int32 Max
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Max");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Max", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual Int32 Min
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Min");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Min", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.Enums.OrientationConstants Orientation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.OrientationConstants>(this, "Orientation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Orientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual bool SelectRange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SelectRange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SelectRange", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual Int32 SelStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SelStart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SelStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual Int32 SelLength
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SelLength");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SelLength", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.Enums.TickStyleConstants TickStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.TickStyleConstants>(this, "TickStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "TickStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual Int32 TickFrequency
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "TickFrequency");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TickFrequency", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual Int32 Value
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Value");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Value", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6), NativeResult]
		public virtual stdole.Picture MouseIcon
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MouseIcon", paramsArray);
                return returnItem as stdole.Picture;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MouseIcon", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.Enums.MousePointerConstants MousePointer
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.MousePointerConstants>(this, "MousePointer");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "MousePointer", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual bool Enabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Enabled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Enabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 hWnd
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "hWnd");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "hWnd", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.Enums.BorderStyleConstants BorderStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.BorderStyleConstants>(this, "BorderStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "BorderStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.Enums.OLEDropConstants OLEDropMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.OLEDropConstants>(this, "OLEDropMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "OLEDropMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual Int32 GetNumTicks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "GetNumTicks");
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string Text
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual NetOffice.MSComctlLibApi.Enums.TextPositionConstants TextPosition
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSComctlLibApi.Enums.TextPositionConstants>(this, "TextPosition");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "TextPosition", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void Refresh()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void ClearSel()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ClearSel");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void DoClick()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DoClick");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void OLEDrag()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OLEDrag");
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void AboutBox()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AboutBox");
		}

		#endregion

		#pragma warning restore
	}
}

