using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSFormsApi;

namespace NetOffice.MSFormsApi.Behind
{
	/// <summary>
	/// DispatchInterface IWHTMLImage 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IWHTMLImage : COMObject, NetOffice.MSFormsApi.IWHTMLImage
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
                    _contractType = typeof(NetOffice.MSFormsApi.IWHTMLImage);
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
                    _type = typeof(IWHTMLImage);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IWHTMLImage() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual string Action
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Action");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Action", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual string Source
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Source");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Source", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual string Encoding
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Encoding");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Encoding", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual string Method
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Method");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Method", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual string HTMLName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HTMLName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HTMLName", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string HTMLType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HTMLType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HTMLType", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}

