using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLOptionElementFactory 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IHTMLOptionElementFactory : COMObject, NetOffice.MSHTMLApi.IHTMLOptionElementFactory
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLOptionElementFactory);
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
                    _type = typeof(IHTMLOptionElementFactory);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLOptionElementFactory() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="text">optional object text</param>
		/// <param name="value">optional object value</param>
		/// <param name="defaultSelected">optional object defaultSelected</param>
		/// <param name="selected">optional object selected</param>
		[SupportByVersion("MSHTML", 4)]
		[BaseResult]
		public virtual NetOffice.MSHTMLApi.IHTMLOptionElement create(object text, object value, object defaultSelected, object selected)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLOptionElement>(this, "create", text, value, defaultSelected, selected);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLOptionElement create()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLOptionElement>(this, "create");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="text">optional object text</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLOptionElement create(object text)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLOptionElement>(this, "create", text);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="text">optional object text</param>
		/// <param name="value">optional object value</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLOptionElement create(object text, object value)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLOptionElement>(this, "create", text, value);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="text">optional object text</param>
		/// <param name="value">optional object value</param>
		/// <param name="defaultSelected">optional object defaultSelected</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("MSHTML", 4)]
		public virtual NetOffice.MSHTMLApi.IHTMLOptionElement create(object text, object value, object defaultSelected)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.MSHTMLApi.IHTMLOptionElement>(this, "create", text, value, defaultSelected);
		}

		#endregion

		#pragma warning restore
	}
}

