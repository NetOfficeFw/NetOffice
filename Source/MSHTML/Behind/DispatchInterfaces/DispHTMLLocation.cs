using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface DispHTMLLocation 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class DispHTMLLocation : COMObject, NetOffice.MSHTMLApi.DispHTMLLocation
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
                    _contractType = typeof(NetOffice.MSHTMLApi.DispHTMLLocation);
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
                    _type = typeof(DispHTMLLocation);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public DispHTMLLocation() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string href
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "href");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "href", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string protocol
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "protocol");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "protocol", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string host
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "host");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "host", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string hostname
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "hostname");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "hostname", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string port
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "port");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "port", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string pathname
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "pathname");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "pathname", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string search
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "search");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "search", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string hash
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "hash");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "hash", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSHTML", 4), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object constructor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "constructor");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="flag">optional bool flag = false</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void reload(object flag)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "reload", flag);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void reload()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "reload");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstr">string bstr</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void replace(string bstr)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "replace", bstr);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstr">string bstr</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void assign(string bstr)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "assign", bstr);
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

