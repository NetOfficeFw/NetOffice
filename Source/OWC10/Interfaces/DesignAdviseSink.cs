using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OWC10Api
{
	///<summary>
	/// Interface DesignAdviseSink 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class DesignAdviseSink : COMObject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		/// 
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		/// <param name="fGrid">Int32 fGrid</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 ObjectAdded(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject, Int32 fGrid)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dscobjtyp, varObject, fGrid);
			object returnItem = Invoker.MethodReturn(this, "ObjectAdded", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 ObjectDeleted(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dscobjtyp, varObject);
			object returnItem = Invoker.MethodReturn(this, "ObjectDeleted", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		/// <param name="bstrRsd">string bstrRsd</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 ObjectMoved(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject, string bstrRsd)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dscobjtyp, varObject, bstrRsd);
			object returnItem = Invoker.MethodReturn(this, "ObjectMoved", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 DataModelLoad()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "DataModelLoad", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 ObjectChanged(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dscobjtyp, varObject);
			object returnItem = Invoker.MethodReturn(this, "ObjectChanged", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 ObjectDeleteComplete(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dscobjtyp);
			object returnItem = Invoker.MethodReturn(this, "ObjectDeleteComplete", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="dscobjtyp">NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp</param>
		/// <param name="varObject">object varObject</param>
		/// <param name="bstrPreviousName">string bstrPreviousName</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 ObjectRenamed(NetOffice.OWC10Api.Enums.DscObjectTypeEnum dscobjtyp, object varObject, string bstrPreviousName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dscobjtyp, varObject, bstrPreviousName);
			object returnItem = Invoker.MethodReturn(this, "ObjectRenamed", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}