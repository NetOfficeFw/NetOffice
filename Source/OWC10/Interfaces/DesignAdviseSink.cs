using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// Interface DesignAdviseSink 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
 	public class DesignAdviseSink : COMObject
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
                    _type = typeof(DesignAdviseSink);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public DesignAdviseSink(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DesignAdviseSink(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DesignAdviseSink(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DesignAdviseSink(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DesignAdviseSink(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DesignAdviseSink(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DesignAdviseSink() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DesignAdviseSink(string progId) : base(progId)
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
		public Int32 ObjectAdded(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject, Int32 fGrid)
		{
			return Factory.ExecuteInt32MethodGet(this, "ObjectAdded", dscobjtyp, varObject, fGrid);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 ObjectDeleted(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject)
		{
			return Factory.ExecuteInt32MethodGet(this, "ObjectDeleted", dscobjtyp, varObject);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		/// <param name="bstrRsd">string bstrRsd</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 ObjectMoved(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject, string bstrRsd)
		{
			return Factory.ExecuteInt32MethodGet(this, "ObjectMoved", dscobjtyp, varObject, bstrRsd);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 DataModelLoad()
		{
			return Factory.ExecuteInt32MethodGet(this, "DataModelLoad");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 ObjectChanged(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject)
		{
			return Factory.ExecuteInt32MethodGet(this, "ObjectChanged", dscobjtyp, varObject);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 ObjectDeleteComplete(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp)
		{
			return Factory.ExecuteInt32MethodGet(this, "ObjectDeleteComplete", dscobjtyp);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		/// <param name="bstrPreviousName">string bstrPreviousName</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 ObjectRenamed(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject, string bstrPreviousName)
		{
			return Factory.ExecuteInt32MethodGet(this, "ObjectRenamed", dscobjtyp, varObject, bstrPreviousName);
		}

		#endregion

		#pragma warning restore
	}
}
