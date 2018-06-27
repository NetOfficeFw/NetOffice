using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLDataTransfer 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLDataTransfer : COMObject, NetOffice.MSHTMLApi.IHTMLDataTransfer
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLDataTransfer);
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
                    _type = typeof(IHTMLDataTransfer);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLDataTransfer() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string dropEffect
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "dropEffect");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "dropEffect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string effectAllowed
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "effectAllowed");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "effectAllowed", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="format">string format</param>
		/// <param name="data">object data</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool setData(string format, object data)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "setData", format, data);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="format">string format</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object getData(string format)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "getData", format);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="format">string format</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool clearData(string format)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "clearData", format);
		}

		#endregion

		#pragma warning restore
	}
}

