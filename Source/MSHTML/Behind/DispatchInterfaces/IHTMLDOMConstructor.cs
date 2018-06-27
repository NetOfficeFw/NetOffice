using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLDOMConstructor 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLDOMConstructor : IHTMLStyle6, NetOffice.MSHTMLApi.IHTMLDOMConstructor
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLDOMConstructor);
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
                    _type = typeof(IHTMLDOMConstructor);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLDOMConstructor() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="propname">string propname</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object LookupGetter(string propname)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LookupGetter", propname);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="propname">string propname</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object LookupSetter(string propname)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LookupSetter", propname);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="propname">string propname</param>
		/// <param name="pdispHandler">object pdispHandler</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void DefineGetter(string propname, object pdispHandler)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DefineGetter", propname, pdispHandler);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="propname">string propname</param>
		/// <param name="pdispHandler">object pdispHandler</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void DefineSetter(string propname, object pdispHandler)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DefineSetter", propname, pdispHandler);
		}

		#endregion

		#pragma warning restore
	}
}

