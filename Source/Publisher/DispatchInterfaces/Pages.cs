using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.PublisherApi
{
	///<summary>
	/// DispatchInterface Pages 
	/// SupportByVersion Publisher, 14,15,16
	///</summary>
	[SupportByVersionAttribute("Publisher", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Pages : COMObject ,IEnumerable<NetOffice.PublisherApi.Page>
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
                    _type = typeof(Pages);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Pages(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="item">Int32 Item</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.PublisherApi.Page this[Int32 item]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(item);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.PublisherApi.Page newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Page.LateBindingApiWrapperType) as NetOffice.PublisherApi.Page;
			return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.PublisherApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Application.LateBindingApiWrapperType) as NetOffice.PublisherApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
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
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="count">Int32 Count</param>
		/// <param name="after">Int32 After</param>
		/// <param name="duplicateObjectsOnPage">optional Int32 DuplicateObjectsOnPage = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Page Add10(Int32 count, Int32 after, object duplicateObjectsOnPage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(count, after, duplicateObjectsOnPage);
			object returnItem = Invoker.MethodReturn(this, "Add10", paramsArray);
			NetOffice.PublisherApi.Page newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Page.LateBindingApiWrapperType) as NetOffice.PublisherApi.Page;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="count">Int32 Count</param>
		/// <param name="after">Int32 After</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Page Add10(Int32 count, Int32 after)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(count, after);
			object returnItem = Invoker.MethodReturn(this, "Add10", paramsArray);
			NetOffice.PublisherApi.Page newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Page.LateBindingApiWrapperType) as NetOffice.PublisherApi.Page;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="after">Int32 After</param>
		/// <param name="pageType">optional NetOffice.PublisherApi.Enums.PbWizardPageType PageType = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void AddWizardPage10(Int32 after, object pageType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(after, pageType);
			Invoker.Method(this, "AddWizardPage10", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="after">Int32 After</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void AddWizardPage10(Int32 after)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(after);
			Invoker.Method(this, "AddWizardPage10", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pageID">Int32 PageID</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Page FindByPageID(Int32 pageID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pageID);
			object returnItem = Invoker.MethodReturn(this, "FindByPageID", paramsArray);
			NetOffice.PublisherApi.Page newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Page.LateBindingApiWrapperType) as NetOffice.PublisherApi.Page;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="count">Int32 Count</param>
		/// <param name="after">Int32 After</param>
		/// <param name="duplicateObjectsOnPage">optional Int32 DuplicateObjectsOnPage = -1</param>
		/// <param name="addHyperlinkToWebNavBar">optional bool AddHyperlinkToWebNavBar = false</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Page Add(Int32 count, Int32 after, object duplicateObjectsOnPage, object addHyperlinkToWebNavBar)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(count, after, duplicateObjectsOnPage, addHyperlinkToWebNavBar);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.PublisherApi.Page newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Page.LateBindingApiWrapperType) as NetOffice.PublisherApi.Page;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="count">Int32 Count</param>
		/// <param name="after">Int32 After</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Page Add(Int32 count, Int32 after)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(count, after);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.PublisherApi.Page newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Page.LateBindingApiWrapperType) as NetOffice.PublisherApi.Page;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="count">Int32 Count</param>
		/// <param name="after">Int32 After</param>
		/// <param name="duplicateObjectsOnPage">optional Int32 DuplicateObjectsOnPage = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Page Add(Int32 count, Int32 after, object duplicateObjectsOnPage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(count, after, duplicateObjectsOnPage);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.PublisherApi.Page newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Page.LateBindingApiWrapperType) as NetOffice.PublisherApi.Page;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="after">Int32 After</param>
		/// <param name="pageType">optional NetOffice.PublisherApi.Enums.PbWizardPageType PageType = -1</param>
		/// <param name="addHyperlinkToWebNavBar">optional bool AddHyperlinkToWebNavBar = false</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void AddWizardPage(Int32 after, object pageType, object addHyperlinkToWebNavBar)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(after, pageType, addHyperlinkToWebNavBar);
			Invoker.Method(this, "AddWizardPage", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="after">Int32 After</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void AddWizardPage(Int32 after)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(after);
			Invoker.Method(this, "AddWizardPage", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="after">Int32 After</param>
		/// <param name="pageType">optional NetOffice.PublisherApi.Enums.PbWizardPageType PageType = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void AddWizardPage(Int32 after, object pageType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(after, pageType);
			Invoker.Method(this, "AddWizardPage", paramsArray);
		}

		#endregion

       #region IEnumerable<NetOffice.PublisherApi.Page> Member
        
        /// <summary>
		/// SupportByVersionAttribute Publisher, 14,15,16
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
       public IEnumerator<NetOffice.PublisherApi.Page> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.PublisherApi.Page item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Publisher, 14,15,16
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}