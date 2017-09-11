using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVServerPublishOptions 
	/// SupportByVersion Visio, 14,15,16
	/// </summary>
	[SupportByVersion("Visio", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IVServerPublishOptions : COMObject
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
                    _type = typeof(IVServerPublishOptions);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IVServerPublishOptions(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IVServerPublishOptions(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVServerPublishOptions(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVServerPublishOptions(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVServerPublishOptions(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVServerPublishOptions(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVServerPublishOptions() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVServerPublishOptions(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int16 Stat
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="pageName">string pageName</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags flags</param>
		[SupportByVersion("Visio", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool get_IsPublishedPage(string pageName, NetOffice.VisioApi.Enums.VisLangFlags flags)
		{
			return Factory.ExecuteBoolPropertyGet(this, "IsPublishedPage", pageName, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Alias for get_IsPublishedPage
		/// </summary>
		/// <param name="pageName">string pageName</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags flags</param>
		[SupportByVersion("Visio", 14,15,16), Redirect("get_IsPublishedPage")]
		public bool IsPublishedPage(string pageName, NetOffice.VisioApi.Enums.VisLangFlags flags)
		{
			return get_IsPublishedPage(pageName, flags);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="pageName">string pageName</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags flags</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void IncludePage(string pageName, NetOffice.VisioApi.Enums.VisLangFlags flags)
		{
			 Factory.ExecuteMethod(this, "IncludePage", pageName, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="pageName">string pageName</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags flags</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void ExcludePage(string pageName, NetOffice.VisioApi.Enums.VisLangFlags flags)
		{
			 Factory.ExecuteMethod(this, "ExcludePage", pageName, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="publishPages">NetOffice.VisioApi.Enums.VisPublishPages publishPages</param>
		/// <param name="namesArray">String[] namesArray</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags flags</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void SetPagesToPublish(NetOffice.VisioApi.Enums.VisPublishPages publishPages, String[] namesArray, NetOffice.VisioApi.Enums.VisLangFlags flags)
		{
            object[] paramsArray = Invoker.ValidateParamsArray(publishPages, (object)namesArray, flags);
            Invoker.Method(this, "SetPagesToPublish", paramsArray);
        }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags flags</param>
		/// <param name="publishPages">NetOffice.VisioApi.Enums.VisPublishPages publishPages</param>
		/// <param name="namesArray">String[] namesArray</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void GetPagesToPublish(NetOffice.VisioApi.Enums.VisLangFlags flags, out NetOffice.VisioApi.Enums.VisPublishPages publishPages, out String[] namesArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true);
			publishPages = 0;
			namesArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray(flags, publishPages, (object)namesArray);
			Invoker.Method(this, "GetPagesToPublish", paramsArray, modifiers);
			publishPages = (NetOffice.VisioApi.Enums.VisPublishPages)paramsArray[1];
			namesArray = (String[])paramsArray[2];
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="publishDataRecordsets">NetOffice.VisioApi.Enums.VisPublishDataRecordsets publishDataRecordsets</param>
		/// <param name="dataRecordsetIDs">Int32[] dataRecordsetIDs</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void SetRecordsetsToPublish(NetOffice.VisioApi.Enums.VisPublishDataRecordsets publishDataRecordsets, Int32[] dataRecordsetIDs)
		{
            object[] paramsArray = Invoker.ValidateParamsArray(publishDataRecordsets, (object)dataRecordsetIDs);
            Invoker.Method(this, "SetRecordsetsToPublish", paramsArray);
        }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="publishDataRecordsets">NetOffice.VisioApi.Enums.VisPublishDataRecordsets publishDataRecordsets</param>
		/// <param name="dataRecordsetIDs">Int32[] dataRecordsetIDs</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void GetRecordsetsToPublish(out NetOffice.VisioApi.Enums.VisPublishDataRecordsets publishDataRecordsets, out Int32[] dataRecordsetIDs)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true);
			publishDataRecordsets = 0;
			dataRecordsetIDs = null;
			object[] paramsArray = Invoker.ValidateParamsArray(publishDataRecordsets, (object)dataRecordsetIDs);
			Invoker.Method(this, "GetRecordsetsToPublish", paramsArray, modifiers);
			publishDataRecordsets = (NetOffice.VisioApi.Enums.VisPublishDataRecordsets)paramsArray[0];
			dataRecordsetIDs = (Int32[])paramsArray[1];
		}

		#endregion

		#pragma warning restore
	}
}
