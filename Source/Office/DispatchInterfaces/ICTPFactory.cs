using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// DispatchInterface ICTPFactory 
	/// SupportByVersion Office, 12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864938.aspx
	///</summary>
	[SupportByVersionAttribute("Office", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class ICTPFactory : COMObject
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
                    _type = typeof(ICTPFactory);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ICTPFactory(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ICTPFactory(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ICTPFactory(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ICTPFactory(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ICTPFactory(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ICTPFactory() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ICTPFactory(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860563.aspx
		/// </summary>
		/// <param name="cTPAxID">string CTPAxID</param>
		/// <param name="cTPTitle">string CTPTitle</param>
		/// <param name="cTPParentWindow">optional object CTPParentWindow</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi._CustomTaskPane CreateCTP(string cTPAxID, string cTPTitle, object cTPParentWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cTPAxID, cTPTitle, cTPParentWindow);
			object returnItem = Invoker.MethodReturn(this, "CreateCTP", paramsArray);
			NetOffice.OfficeApi._CustomTaskPane newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OfficeApi._CustomTaskPane;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860563.aspx
		/// </summary>
		/// <param name="cTPAxID">string CTPAxID</param>
		/// <param name="cTPTitle">string CTPTitle</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi._CustomTaskPane CreateCTP(string cTPAxID, string cTPTitle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cTPAxID, cTPTitle);
			object returnItem = Invoker.MethodReturn(this, "CreateCTP", paramsArray);
			NetOffice.OfficeApi._CustomTaskPane newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OfficeApi._CustomTaskPane;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}