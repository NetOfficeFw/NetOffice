using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLDialog 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IHTMLDialog : COMObject, NetOffice.MSHTMLApi.IHTMLDialog
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLDialog);
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
                    _type = typeof(IHTMLDialog);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLDialog() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object dialogTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "dialogTop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "dialogTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object dialogLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "dialogLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "dialogLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object dialogWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "dialogWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "dialogWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object dialogHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "dialogHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "dialogHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object dialogArguments
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "dialogArguments");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object menuArguments
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "menuArguments");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object returnValue
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "returnValue");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "returnValue", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void close()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "close");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string toString()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "toString");
		}

		#endregion

		#pragma warning restore
	}
}

