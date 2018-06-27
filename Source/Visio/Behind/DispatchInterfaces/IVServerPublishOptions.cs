using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.VisioApi;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// DispatchInterface IVServerPublishOptions 
	/// SupportByVersion Visio, 14,15,16
	/// </summary>
	[SupportByVersion("Visio", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IVServerPublishOptions : COMObject, NetOffice.VisioApi.IVServerPublishOptions
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
                    _contractType = typeof(NetOffice.VisioApi.IVServerPublishOptions);
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
                    _type = typeof(IVServerPublishOptions);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IVServerPublishOptions() : base()
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
		public virtual NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int16 Stat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual Int16 ObjectType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ObjectType");
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
		public virtual bool get_IsPublishedPage(string pageName, NetOffice.VisioApi.Enums.VisLangFlags flags)
		{
			return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsPublishedPage", pageName, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Alias for get_IsPublishedPage
		/// </summary>
		/// <param name="pageName">string pageName</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags flags</param>
		[SupportByVersion("Visio", 14,15,16), Redirect("get_IsPublishedPage")]
		public virtual bool IsPublishedPage(string pageName, NetOffice.VisioApi.Enums.VisLangFlags flags)
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
		public virtual void IncludePage(string pageName, NetOffice.VisioApi.Enums.VisLangFlags flags)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "IncludePage", pageName, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="pageName">string pageName</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags flags</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void ExcludePage(string pageName, NetOffice.VisioApi.Enums.VisLangFlags flags)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExcludePage", pageName, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="publishPages">NetOffice.VisioApi.Enums.VisPublishPages publishPages</param>
		/// <param name="namesArray">String[] namesArray</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags flags</param>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual void SetPagesToPublish(NetOffice.VisioApi.Enums.VisPublishPages publishPages, String[] namesArray, NetOffice.VisioApi.Enums.VisLangFlags flags)
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
		public virtual void GetPagesToPublish(NetOffice.VisioApi.Enums.VisLangFlags flags, out NetOffice.VisioApi.Enums.VisPublishPages publishPages, out String[] namesArray)
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
		public virtual void SetRecordsetsToPublish(NetOffice.VisioApi.Enums.VisPublishDataRecordsets publishDataRecordsets, Int32[] dataRecordsetIDs)
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
		public virtual void GetRecordsetsToPublish(out NetOffice.VisioApi.Enums.VisPublishDataRecordsets publishDataRecordsets, out Int32[] dataRecordsetIDs)
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

