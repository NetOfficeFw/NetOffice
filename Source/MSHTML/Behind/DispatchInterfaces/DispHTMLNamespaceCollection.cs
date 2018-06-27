using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface DispHTMLNamespaceCollection 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class DispHTMLNamespaceCollection : COMObject, NetOffice.MSHTMLApi.DispHTMLNamespaceCollection
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
                    _contractType = typeof(NetOffice.MSHTMLApi.DispHTMLNamespaceCollection);
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
                    _type = typeof(DispHTMLNamespaceCollection);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public DispHTMLNamespaceCollection() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 length
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "length");
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
		/// <param name="index">object index</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object item(object index)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "item", index);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrNamespace">string bstrNamespace</param>
		/// <param name="bstrUrn">string bstrUrn</param>
		/// <param name="implementationUrl">optional object implementationUrl</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object add(string bstrNamespace, string bstrUrn, object implementationUrl)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "add", bstrNamespace, bstrUrn, implementationUrl);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrNamespace">string bstrNamespace</param>
		/// <param name="bstrUrn">string bstrUrn</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual object add(string bstrNamespace, string bstrUrn)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "add", bstrNamespace, bstrUrn);
		}

		#endregion

		#pragma warning restore
	}
}

