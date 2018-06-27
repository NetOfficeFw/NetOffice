using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.VisioApi;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// Interface IVisLibOcxSupport 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class IVisLibOcxSupport : COMObject, NetOffice.VisioApi.IVisLibOcxSupport
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
                    _contractType = typeof(NetOffice.VisioApi.IVisLibOcxSupport);
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
                    _type = typeof(IVisLibOcxSupport);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IVisLibOcxSupport() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string HostID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HostID");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HostID", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVWindow Window
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVWindow>(this, "Window");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string Src
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Src");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Src", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool NegotiateMenus
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "NegotiateMenus");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NegotiateMenus", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual bool NegotiateToolbars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "NegotiateToolbars");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NegotiateToolbars", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.Enums.VisPageSizingBehaviors PageSizingBehavior
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisPageSizingBehaviors>(this, "PageSizingBehavior");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PageSizingBehavior", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 About()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "About");
		}

		#endregion

		#pragma warning restore
	}
}

