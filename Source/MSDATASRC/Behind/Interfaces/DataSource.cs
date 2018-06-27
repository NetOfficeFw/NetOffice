using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSDATASRCApi;

namespace NetOffice.MSDATASRCApi.Behind
{
	/// <summary>
	/// Interface DataSource 
	/// SupportByVersion MSDATASRC, 4
	/// </summary>
	[SupportByVersion("MSDATASRC", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class DataSource : COMObject, NetOffice.MSDATASRCApi.DataSource
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
                    _contractType = typeof(NetOffice.MSDATASRCApi.DataSource);
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
                    _type = typeof(DataSource);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor
		
		/// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
		public DataSource() : base()
		{
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		/// <param name="bstrDM">string bstrDM</param>
		/// <param name="riid">Guid riid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		public virtual object getDataMember(string bstrDM, Guid riid)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "getDataMember", bstrDM, riid);
		}

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		/// <param name="lIndex">Int32 lIndex</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		public virtual string getDataMemberName(Int32 lIndex)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "getDataMemberName", lIndex);
		}

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		public virtual Int32 getDataMemberCount()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "getDataMemberCount");
		}

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		/// <param name="pDSL">NetOffice.MSDATASRCApi.DataSourceListener pDSL</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		public virtual Int32 addDataSourceListener(NetOffice.MSDATASRCApi.DataSourceListener pDSL)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "addDataSourceListener", pDSL);
		}

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		/// <param name="pDSL">NetOffice.MSDATASRCApi.DataSourceListener pDSL</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		public virtual Int32 removeDataSourceListener(NetOffice.MSDATASRCApi.DataSourceListener pDSL)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "removeDataSourceListener", pDSL);
		}

		#endregion

		#pragma warning restore
	}
}
