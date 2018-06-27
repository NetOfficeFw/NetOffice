using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OutlookApi;

namespace NetOffice.OutlookApi.Behind
{
	/// <summary>
	/// DispatchInterface _DInspectorCtrl 
	/// SupportByVersion Outlook, 10
	/// </summary>
	[SupportByVersion("Outlook", 10)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _DInspectorCtrl : COMObject, NetOffice.OutlookApi._DInspectorCtrl
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
                    _contractType = typeof(NetOffice.OutlookApi._DInspectorCtrl);
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
                    _type = typeof(_DInspectorCtrl);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _DInspectorCtrl() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 10
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 10)]
		public virtual string URL
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "URL");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "URL", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 10), ProxyResult]
		public virtual object Item
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Item");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 10
		/// </summary>
		/// <param name="pdispItem">object pdispItem</param>
		[SupportByVersion("Outlook", 10)]
		public virtual void OnItemChange(object pdispItem)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OnItemChange", pdispItem);
		}

		#endregion

		#pragma warning restore
	}
}

