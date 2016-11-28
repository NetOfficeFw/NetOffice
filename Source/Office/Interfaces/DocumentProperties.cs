using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// Interface DocumentProperties 
	/// SupportByVersion Office, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863102.aspx
	///</summary>
	[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class DocumentProperties : COMObject ,IEnumerable<NetOffice.OfficeApi.DocumentProperty>
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
                    _type = typeof(DocumentProperties);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DocumentProperties(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DocumentProperties(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DocumentProperties(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DocumentProperties(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DocumentProperties(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DocumentProperties() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DocumentProperties(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864976.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
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

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.OfficeApi.DocumentProperty this[object index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.OfficeApi.DocumentProperty newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.DocumentProperty.LateBindingApiWrapperType) as NetOffice.OfficeApi.DocumentProperty;
			return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863076.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864029.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public object Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861448.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862806.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="linkToContent">bool LinkToContent</param>
		/// <param name="type">optional object Type</param>
		/// <param name="value">optional object Value</param>
		/// <param name="linkSource">optional object LinkSource</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.DocumentProperty Add(string name, bool linkToContent, object type, object value, object linkSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, linkToContent, type, value, linkSource);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.DocumentProperty newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.DocumentProperty.LateBindingApiWrapperType) as NetOffice.OfficeApi.DocumentProperty;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862806.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="linkToContent">bool LinkToContent</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.DocumentProperty Add(string name, bool linkToContent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, linkToContent);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.DocumentProperty newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.DocumentProperty.LateBindingApiWrapperType) as NetOffice.OfficeApi.DocumentProperty;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862806.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="linkToContent">bool LinkToContent</param>
		/// <param name="type">optional object Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.DocumentProperty Add(string name, bool linkToContent, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, linkToContent, type);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.DocumentProperty newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.DocumentProperty.LateBindingApiWrapperType) as NetOffice.OfficeApi.DocumentProperty;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862806.aspx
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="linkToContent">bool LinkToContent</param>
		/// <param name="type">optional object Type</param>
		/// <param name="value">optional object Value</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.DocumentProperty Add(string name, bool linkToContent, object type, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, linkToContent, type, value);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.DocumentProperty newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.DocumentProperty.LateBindingApiWrapperType) as NetOffice.OfficeApi.DocumentProperty;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.OfficeApi.DocumentProperty> Member
        
        /// <summary>
		/// SupportByVersionAttribute Office, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
       public IEnumerator<NetOffice.OfficeApi.DocumentProperty> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.OfficeApi.DocumentProperty item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Office, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}