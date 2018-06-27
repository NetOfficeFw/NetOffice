using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLBodyElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IHTMLBodyElement : IHTMLTextContainer, NetOffice.MSHTMLApi.IHTMLBodyElement
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLBodyElement);
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
                    _type = typeof(IHTMLBodyElement);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLBodyElement() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string background
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "background");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "background", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string bgProperties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "bgProperties");			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "bgProperties", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object leftMargin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "leftMargin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "leftMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object topMargin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "topMargin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "topMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object rightMargin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "rightMargin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "rightMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object bottomMargin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "bottomMargin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "bottomMargin", value);
			}
		}

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
		public virtual object text
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "text");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "text", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object link
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "link");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "link", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object vLink
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "vLink");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "vLink", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object aLink
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "aLink");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "aLink", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onload
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onload");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onload", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onunload
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onunload");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onunload", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string scroll
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "scroll");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "scroll", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onselect
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onselect");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onselect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onbeforeunload
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onbeforeunload");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onbeforeunload", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLTxtRange createTextRange()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLTxtRange>(this, "createTextRange", typeof(NetOffice.MSHTMLApi.IHTMLTxtRange));
		}

		#endregion

		#pragma warning restore
	}
}


