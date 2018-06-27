using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLScriptElement 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IHTMLScriptElement : COMObject, NetOffice.MSHTMLApi.IHTMLScriptElement
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLScriptElement);
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
                    _type = typeof(IHTMLScriptElement);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLScriptElement() : base()
		{

		}

		#endregion
		
		#region Properties

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
		public virtual string htmlFor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "htmlFor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "htmlFor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string get_event()
		{
			return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "event");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual void set_event(string value)
		{
			InvokerService.InvokeInternal.ExecutePropertySet(this, "event", value);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string text
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "text");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "text", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool defer
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "defer");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "defer", value);
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

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}

