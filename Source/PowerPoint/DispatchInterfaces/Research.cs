using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.PowerPointApi
{
	///<summary>
	/// DispatchInterface Research 
	/// SupportByVersion PowerPoint, 12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745646.aspx
	///</summary>
	[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Research : COMObject
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
                    _type = typeof(Research);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Research(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Research(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Research(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Research(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Research(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Research() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Research(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746098.aspx
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.PowerPointApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Application.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744070.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744220.aspx
		/// </summary>
		/// <param name="serviceID">string ServiceID</param>
		/// <param name="queryString">optional object QueryString</param>
		/// <param name="queryLanguage">optional object QueryLanguage</param>
		/// <param name="useSelection">optional bool UseSelection = false</param>
		/// <param name="launchQuery">optional bool LaunchQuery = true</param>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void Query(string serviceID, object queryString, object queryLanguage, object useSelection, object launchQuery)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(serviceID, queryString, queryLanguage, useSelection, launchQuery);
			Invoker.Method(this, "Query", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744220.aspx
		/// </summary>
		/// <param name="serviceID">string ServiceID</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void Query(string serviceID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(serviceID);
			Invoker.Method(this, "Query", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744220.aspx
		/// </summary>
		/// <param name="serviceID">string ServiceID</param>
		/// <param name="queryString">optional object QueryString</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void Query(string serviceID, object queryString)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(serviceID, queryString);
			Invoker.Method(this, "Query", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744220.aspx
		/// </summary>
		/// <param name="serviceID">string ServiceID</param>
		/// <param name="queryString">optional object QueryString</param>
		/// <param name="queryLanguage">optional object QueryLanguage</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void Query(string serviceID, object queryString, object queryLanguage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(serviceID, queryString, queryLanguage);
			Invoker.Method(this, "Query", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744220.aspx
		/// </summary>
		/// <param name="serviceID">string ServiceID</param>
		/// <param name="queryString">optional object QueryString</param>
		/// <param name="queryLanguage">optional object QueryLanguage</param>
		/// <param name="useSelection">optional bool UseSelection = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void Query(string serviceID, object queryString, object queryLanguage, object useSelection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(serviceID, queryString, queryLanguage, useSelection);
			Invoker.Method(this, "Query", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745349.aspx
		/// </summary>
		/// <param name="language1">object Language1</param>
		/// <param name="language2">object Language2</param>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public void SetLanguagePair(object language1, object language2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(language1, language2);
			Invoker.Method(this, "SetLanguagePair", paramsArray);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746351.aspx
		/// </summary>
		/// <param name="serviceID">string ServiceID</param>
		[SupportByVersionAttribute("PowerPoint", 12,14,15,16)]
		public bool IsResearchService(string serviceID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(serviceID);
			object returnItem = Invoker.MethodReturn(this, "IsResearchService", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}