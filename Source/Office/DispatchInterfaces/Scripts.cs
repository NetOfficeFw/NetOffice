using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// DispatchInterface Scripts 
	/// SupportByVersion Office, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Scripts : _IMsoDispObj ,IEnumerable<NetOffice.OfficeApi.Script>
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
                    _type = typeof(Scripts);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Scripts(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Scripts(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Scripts(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Scripts(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Scripts(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Scripts() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Scripts(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.OfficeApi.Script this[object index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.OfficeApi.Script newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Script.LateBindingApiWrapperType) as NetOffice.OfficeApi.Script;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
		/// <param name="location">optional NetOffice.OfficeApi.Enums.MsoScriptLocation Location = 2</param>
		/// <param name="language">optional NetOffice.OfficeApi.Enums.MsoScriptLanguage Language = 2</param>
		/// <param name="id">optional string Id = </param>
		/// <param name="extended">optional string Extended = </param>
		/// <param name="scriptText">optional string ScriptText = </param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Script Add(object anchor, object location, object language, object id, object extended, object scriptText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(anchor, location, language, id, extended, scriptText);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.Script newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Script.LateBindingApiWrapperType) as NetOffice.OfficeApi.Script;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Script Add()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.Script newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Script.LateBindingApiWrapperType) as NetOffice.OfficeApi.Script;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Script Add(object anchor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(anchor);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.Script newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Script.LateBindingApiWrapperType) as NetOffice.OfficeApi.Script;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
		/// <param name="location">optional NetOffice.OfficeApi.Enums.MsoScriptLocation Location = 2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Script Add(object anchor, object location)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(anchor, location);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.Script newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Script.LateBindingApiWrapperType) as NetOffice.OfficeApi.Script;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
		/// <param name="location">optional NetOffice.OfficeApi.Enums.MsoScriptLocation Location = 2</param>
		/// <param name="language">optional NetOffice.OfficeApi.Enums.MsoScriptLanguage Language = 2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Script Add(object anchor, object location, object language)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(anchor, location, language);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.Script newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Script.LateBindingApiWrapperType) as NetOffice.OfficeApi.Script;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
		/// <param name="location">optional NetOffice.OfficeApi.Enums.MsoScriptLocation Location = 2</param>
		/// <param name="language">optional NetOffice.OfficeApi.Enums.MsoScriptLanguage Language = 2</param>
		/// <param name="id">optional string Id = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Script Add(object anchor, object location, object language, object id)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(anchor, location, language, id);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.Script newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Script.LateBindingApiWrapperType) as NetOffice.OfficeApi.Script;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
		/// <param name="location">optional NetOffice.OfficeApi.Enums.MsoScriptLocation Location = 2</param>
		/// <param name="language">optional NetOffice.OfficeApi.Enums.MsoScriptLanguage Language = 2</param>
		/// <param name="id">optional string Id = </param>
		/// <param name="extended">optional string Extended = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Script Add(object anchor, object location, object language, object id, object extended)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(anchor, location, language, id, extended);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.Script newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Script.LateBindingApiWrapperType) as NetOffice.OfficeApi.Script;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		#endregion

       #region IEnumerable<NetOffice.OfficeApi.Script> Member
        
        /// <summary>
		/// SupportByVersionAttribute Office, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
       public IEnumerator<NetOffice.OfficeApi.Script> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.OfficeApi.Script item in innerEnumerator)
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