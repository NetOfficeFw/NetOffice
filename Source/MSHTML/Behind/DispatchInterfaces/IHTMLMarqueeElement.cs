using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLMarqueeElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLMarqueeElement : COMObject, NetOffice.MSHTMLApi.IHTMLMarqueeElement
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLMarqueeElement);
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
                    _type = typeof(IHTMLMarqueeElement);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLMarqueeElement() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object bgColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "bgColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "bgColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 scrollDelay
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "scrollDelay");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "scrollDelay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string direction
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "direction");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "direction", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string behavior
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "behavior");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "behavior", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 scrollAmount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "scrollAmount");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "scrollAmount", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 loop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "loop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "loop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 vspace
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "vspace");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "vspace", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 hspace
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "hspace");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "hspace", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onfinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onfinish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onfinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onstart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onstart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onstart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onbounce
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onbounce");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onbounce", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object width
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "width");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "width", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object height
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "height");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "height", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool trueSpeed
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "trueSpeed");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "trueSpeed", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void Start()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Start");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void stop()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "stop");
		}

		#endregion

		#pragma warning restore
	}
}

