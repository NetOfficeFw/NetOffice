using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLAnchorElement2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IHTMLAnchorElement2 : IHTMLAnchorElement, NetOffice.MSHTMLApi.IHTMLAnchorElement2
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLAnchorElement2);
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
                    _type = typeof(IHTMLAnchorElement2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLAnchorElement2() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string charset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "charset");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "charset", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string coords
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "coords");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "coords", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string hreflang
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "hreflang");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "hreflang", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string shape
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "shape");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "shape", value);
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

