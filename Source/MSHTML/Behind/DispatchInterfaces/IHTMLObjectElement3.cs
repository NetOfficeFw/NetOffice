using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLObjectElement3 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLObjectElement3 : IHTMLObjectElement2, NetOffice.MSHTMLApi.IHTMLObjectElement3
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLObjectElement3);
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
                    _type = typeof(IHTMLObjectElement3);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLObjectElement3() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string archive
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "archive");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "archive", value);
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
		public virtual bool declare
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "declare");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "declare", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string standby
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "standby");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "standby", value);
			}
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
		public virtual string useMap
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "useMap");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "useMap", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}

