using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface EmailSignature 
	/// SupportByVersion Word, 9,10,11,12,14
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class EmailSignature : COMObject
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
                    _type = typeof(EmailSignature);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public EmailSignature(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public EmailSignature(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public EmailSignature(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public EmailSignature() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public EmailSignature(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public string NewMessageSignature
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NewMessageSignature", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "NewMessageSignature", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public string ReplyMessageSignature
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReplyMessageSignature", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ReplyMessageSignature", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14)]
		public NetOffice.WordApi.EmailSignatureEntries EmailSignatureEntries
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EmailSignatureEntries", paramsArray);
				NetOffice.WordApi.EmailSignatureEntries newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.EmailSignatureEntries.LateBindingApiWrapperType) as NetOffice.WordApi.EmailSignatureEntries;
				return newObject;
			}
		}

		#endregion

		#region Methods

		#endregion
		#pragma warning restore
	}
}