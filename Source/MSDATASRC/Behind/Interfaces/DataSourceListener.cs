using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSDATASRCApi;

namespace NetOffice.MSDATASRCApi.Behind
{
	/// <summary>
	/// Interface DataSourceListener 
	/// SupportByVersion MSDATASRC, 4
	/// </summary>
	[SupportByVersion("MSDATASRC", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class DataSourceListener : COMObject, NetOffice.MSDATASRCApi.DataSourceListener
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
                    _contractType = typeof(NetOffice.MSDATASRCApi.DataSourceListener);
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
                    _type = typeof(DataSourceListener);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
		public DataSourceListener() : base()
		{
		}		
		
		#endregion
		
		#region Methods

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		/// <param name="bstrDM">string bstrDM</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		public virtual Int32 dataMemberChanged(string bstrDM)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "dataMemberChanged", bstrDM);
		}

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		/// <param name="bstrDM">string bstrDM</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		public virtual Int32 dataMemberAdded(string bstrDM)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "dataMemberAdded", bstrDM);
		}

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		/// <param name="bstrDM">string bstrDM</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		public virtual Int32 dataMemberRemoved(string bstrDM)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "dataMemberRemoved", bstrDM);
		}

		#endregion

		#pragma warning restore
	}
}
