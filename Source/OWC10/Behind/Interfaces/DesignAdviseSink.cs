using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// Interface DesignAdviseSink 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
 	public class DesignAdviseSink : COMObject, NetOffice.OWC10Api.DesignAdviseSink
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
                    _contractType = typeof(NetOffice.OWC10Api.DesignAdviseSink);
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
                    _type = typeof(DesignAdviseSink);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public DesignAdviseSink() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		/// <param name="fGrid">Int32 fGrid</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 ObjectAdded(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject, Int32 fGrid)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ObjectAdded", dscobjtyp, varObject, fGrid);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 ObjectDeleted(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ObjectDeleted", dscobjtyp, varObject);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		/// <param name="bstrRsd">string bstrRsd</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 ObjectMoved(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject, string bstrRsd)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ObjectMoved", dscobjtyp, varObject, bstrRsd);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 DataModelLoad()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DataModelLoad");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 ObjectChanged(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ObjectChanged", dscobjtyp, varObject);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 ObjectDeleteComplete(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ObjectDeleteComplete", dscobjtyp);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		/// <param name="bstrPreviousName">string bstrPreviousName</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 ObjectRenamed(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject, string bstrPreviousName)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ObjectRenamed", dscobjtyp, varObject, bstrPreviousName);
		}

		#endregion

		#pragma warning restore
	}
}
