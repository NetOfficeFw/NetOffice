using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLInputElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IHTMLInputElement : COMObject, NetOffice.MSHTMLApi.IHTMLInputElement
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLInputElement);
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
                    _type = typeof(IHTMLInputElement);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLInputElement() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "type");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "type", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string value
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "value");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "value", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "name");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "name", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool status
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "status");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "status", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool disabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "disabled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "disabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLFormElement form
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSHTMLApi.IHTMLFormElement>(this, "form");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 size
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "size");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "size", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 maxLength
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "maxLength");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "maxLength", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onchange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onchange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onchange", value);
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string defaultValue
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "defaultValue");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "defaultValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool readOnly
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "readOnly");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "readOnly", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool indeterminate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "indeterminate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "indeterminate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool defaultChecked
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "defaultChecked");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "defaultChecked", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool get_checked()
		{
			return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "checked");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual void set_checked(bool value)
		{
			InvokerService.InvokeInternal.ExecutePropertySet(this, "checked", value);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object border
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "border");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "border", value);
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
		public virtual string alt
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "alt");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "alt", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string src
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "src");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "src", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string lowsrc
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "lowsrc");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "lowsrc", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string vrml
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "vrml");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "vrml", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string dynsrc
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "dynsrc");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "dynsrc", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string readyState
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "readyState");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool complete
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "complete");
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object loop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "loop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "loop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string align
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "align");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "align", value);
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
		public virtual object onerror
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onerror");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onerror", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual object onabort
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "onabort");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "onabort", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 width
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "width");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "width", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 height
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "height");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "height", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string Start
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Start");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Start", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void select()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "select");
		}

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


