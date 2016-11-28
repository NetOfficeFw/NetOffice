using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OWC10Api
{
	///<summary>
	/// DispatchInterface Section 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Section : COMObject
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
                    _type = typeof(Section);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Section(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Section(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Section(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Section(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Section(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Section() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Section(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.SectTypeEnum Type
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Type", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OWC10Api.Enums.SectTypeEnum)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.DataPage DataPage
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DataPage", paramsArray);
				NetOffice.OWC10Api.DataPage newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.DataPage.LateBindingApiWrapperType) as NetOffice.OWC10Api.DataPage;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public object HTMLContainer
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HTMLContainer", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Section ParentSection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ParentSection", paramsArray);
				NetOffice.OWC10Api.Section newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Section.LateBindingApiWrapperType) as NetOffice.OWC10Api.Section;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Section ChildSection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChildSection", paramsArray);
				NetOffice.OWC10Api.Section newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Section.LateBindingApiWrapperType) as NetOffice.OWC10Api.Section;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Section NextSibling
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NextSibling", paramsArray);
				NetOffice.OWC10Api.Section newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Section.LateBindingApiWrapperType) as NetOffice.OWC10Api.Section;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Section PreviousSibling
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PreviousSibling", paramsArray);
				NetOffice.OWC10Api.Section newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Section.LateBindingApiWrapperType) as NetOffice.OWC10Api.Section;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Section NextSection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NextSection", paramsArray);
				NetOffice.OWC10Api.Section newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Section.LateBindingApiWrapperType) as NetOffice.OWC10Api.Section;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Section PreviousSection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PreviousSection", paramsArray);
				NetOffice.OWC10Api.Section newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Section.LateBindingApiWrapperType) as NetOffice.OWC10Api.Section;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public bool IsExpanded
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsExpanded", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Expand()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Expand", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void Collapse()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Collapse", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="scrollIntoView">optional bool ScrollIntoView = false</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void MakeCurrent(object scrollIntoView)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(scrollIntoView);
			Invoker.Method(this, "MakeCurrent", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void MakeCurrent()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "MakeCurrent", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}