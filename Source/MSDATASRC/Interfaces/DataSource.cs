using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSDATASRCApi
{
	/// <summary>
	/// Interface DataSource 
	/// SupportByVersion MSDATASRC, 4
	/// </summary>
	[SupportByVersion("MSDATASRC", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class DataSource : COMObject
	{
		#pragma warning disable

		#region Type Information

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

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public DataSource(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DataSource(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DataSource(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DataSource(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DataSource(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DataSource(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DataSource() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DataSource(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		/// <param name="bstrDM">string bstrDM</param>
		/// <param name="riid">Guid riid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		public object getDataMember(string bstrDM, Guid riid)
		{
			return Factory.ExecuteVariantMethodGet(this, "getDataMember", bstrDM, riid);
		}

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		/// <param name="lIndex">Int32 lIndex</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		public string getDataMemberName(Int32 lIndex)
		{
			return Factory.ExecuteStringMethodGet(this, "getDataMemberName", lIndex);
		}

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		public Int32 getDataMemberCount()
		{
			return Factory.ExecuteInt32MethodGet(this, "getDataMemberCount");
		}

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		/// <param name="pDSL">NetOffice.MSDATASRCApi.DataSourceListener pDSL</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		public Int32 addDataSourceListener(NetOffice.MSDATASRCApi.DataSourceListener pDSL)
		{
			return Factory.ExecuteInt32MethodGet(this, "addDataSourceListener", pDSL);
		}

		/// <summary>
		/// SupportByVersion MSDATASRC 4
		/// </summary>
		/// <param name="pDSL">NetOffice.MSDATASRCApi.DataSourceListener pDSL</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSDATASRC", 4)]
		public Int32 removeDataSourceListener(NetOffice.MSDATASRCApi.DataSourceListener pDSL)
		{
			return Factory.ExecuteInt32MethodGet(this, "removeDataSourceListener", pDSL);
		}

		#endregion

		#pragma warning restore
	}
}
