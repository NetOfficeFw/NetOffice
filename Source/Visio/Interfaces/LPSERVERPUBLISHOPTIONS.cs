using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.VisioApi
{
	///<summary>
	/// Interface LPSERVERPUBLISHOPTIONS 
	/// SupportByVersion Visio, 14,15,16
	///</summary>
	[SupportByVersionAttribute("Visio", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class LPSERVERPUBLISHOPTIONS : COMObject
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
                    _type = typeof(LPSERVERPUBLISHOPTIONS);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPSERVERPUBLISHOPTIONS(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPSERVERPUBLISHOPTIONS(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPSERVERPUBLISHOPTIONS(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPSERVERPUBLISHOPTIONS(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPSERVERPUBLISHOPTIONS(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPSERVERPUBLISHOPTIONS() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPSERVERPUBLISHOPTIONS(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.VisioApi.IVApplication newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVApplication;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int16 Stat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Stat", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Document", paramsArray);
				NetOffice.VisioApi.IVDocument newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVDocument;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ObjectType", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="pageName">string PageName</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags Flags</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool get_IsPublishedPage(string pageName, NetOffice.VisioApi.Enums.VisLangFlags flags)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(pageName, flags);
			object returnItem = Invoker.PropertyGet(this, "IsPublishedPage", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Alias for get_IsPublishedPage
		/// </summary>
		/// <param name="pageName">string PageName</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags Flags</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public bool IsPublishedPage(string pageName, NetOffice.VisioApi.Enums.VisLangFlags flags)
		{
			return get_IsPublishedPage(pageName, flags);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pageName">string PageName</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags Flags</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void IncludePage(string pageName, NetOffice.VisioApi.Enums.VisLangFlags flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pageName, flags);
			Invoker.Method(this, "IncludePage", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pageName">string PageName</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags Flags</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void ExcludePage(string pageName, NetOffice.VisioApi.Enums.VisLangFlags flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pageName, flags);
			Invoker.Method(this, "ExcludePage", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="publishPages">NetOffice.VisioApi.Enums.VisPublishPages PublishPages</param>
		/// <param name="namesArray">String[] NamesArray</param>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags Flags</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void SetPagesToPublish(NetOffice.VisioApi.Enums.VisPublishPages publishPages, String[] namesArray, NetOffice.VisioApi.Enums.VisLangFlags flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(publishPages, (object)namesArray, flags);
			Invoker.Method(this, "SetPagesToPublish", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisLangFlags Flags</param>
		/// <param name="publishPages">NetOffice.VisioApi.Enums.VisPublishPages PublishPages</param>
		/// <param name="namesArray">String[] NamesArray</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
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
		/// 
		/// </summary>
		/// <param name="publishDataRecordsets">NetOffice.VisioApi.Enums.VisPublishDataRecordsets PublishDataRecordsets</param>
		/// <param name="dataRecordsetIDs">Int32[] DataRecordsetIDs</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void SetRecordsetsToPublish(NetOffice.VisioApi.Enums.VisPublishDataRecordsets publishDataRecordsets, Int32[] dataRecordsetIDs)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(publishDataRecordsets, (object)dataRecordsetIDs);
			Invoker.Method(this, "SetRecordsetsToPublish", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="publishDataRecordsets">NetOffice.VisioApi.Enums.VisPublishDataRecordsets PublishDataRecordsets</param>
		/// <param name="dataRecordsetIDs">Int32[] DataRecordsetIDs</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
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