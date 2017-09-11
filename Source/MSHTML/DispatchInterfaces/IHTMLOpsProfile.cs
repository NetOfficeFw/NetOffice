using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// DispatchInterface IHTMLOpsProfile 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IHTMLOpsProfile : COMObject
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
                    _type = typeof(IHTMLOpsProfile);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IHTMLOpsProfile(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLOpsProfile(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLOpsProfile(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLOpsProfile(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLOpsProfile(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLOpsProfile(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLOpsProfile() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLOpsProfile(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="reserved">optional object reserved</param>
		[SupportByVersion("MSHTML", 4)]
		public bool addRequest(string name, object reserved)
		{
			return Factory.ExecuteBoolMethodGet(this, "addRequest", name, reserved);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public bool addRequest(string name)
		{
			return Factory.ExecuteBoolMethodGet(this, "addRequest", name);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public void clearRequest()
		{
			 Factory.ExecuteMethod(this, "clearRequest");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		/// <param name="path">optional object path</param>
		/// <param name="expire">optional object expire</param>
		/// <param name="reserved">optional object reserved</param>
		[SupportByVersion("MSHTML", 4)]
		public void doRequest(object usage, object fname, object domain, object path, object expire, object reserved)
		{
			 Factory.ExecuteMethod(this, "doRequest", new object[]{ usage, fname, domain, path, expire, reserved });
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void doRequest(object usage)
		{
			 Factory.ExecuteMethod(this, "doRequest", usage);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void doRequest(object usage, object fname)
		{
			 Factory.ExecuteMethod(this, "doRequest", usage, fname);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void doRequest(object usage, object fname, object domain)
		{
			 Factory.ExecuteMethod(this, "doRequest", usage, fname, domain);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		/// <param name="path">optional object path</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void doRequest(object usage, object fname, object domain, object path)
		{
			 Factory.ExecuteMethod(this, "doRequest", usage, fname, domain, path);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		/// <param name="path">optional object path</param>
		/// <param name="expire">optional object expire</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void doRequest(object usage, object fname, object domain, object path, object expire)
		{
			 Factory.ExecuteMethod(this, "doRequest", new object[]{ usage, fname, domain, path, expire });
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("MSHTML", 4)]
		public string getAttribute(string name)
		{
			return Factory.ExecuteStringMethodGet(this, "getAttribute", name);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="value">string value</param>
		/// <param name="prefs">optional object prefs</param>
		[SupportByVersion("MSHTML", 4)]
		public bool setAttribute(string name, string value, object prefs)
		{
			return Factory.ExecuteBoolMethodGet(this, "setAttribute", name, value, prefs);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="value">string value</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public bool setAttribute(string name, string value)
		{
			return Factory.ExecuteBoolMethodGet(this, "setAttribute", name, value);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool commitChanges()
		{
			return Factory.ExecuteBoolMethodGet(this, "commitChanges");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="reserved">optional object reserved</param>
		[SupportByVersion("MSHTML", 4)]
		public bool addReadRequest(string name, object reserved)
		{
			return Factory.ExecuteBoolMethodGet(this, "addReadRequest", name, reserved);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public bool addReadRequest(string name)
		{
			return Factory.ExecuteBoolMethodGet(this, "addReadRequest", name);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		/// <param name="path">optional object path</param>
		/// <param name="expire">optional object expire</param>
		/// <param name="reserved">optional object reserved</param>
		[SupportByVersion("MSHTML", 4)]
		public void doReadRequest(object usage, object fname, object domain, object path, object expire, object reserved)
		{
			 Factory.ExecuteMethod(this, "doReadRequest", new object[]{ usage, fname, domain, path, expire, reserved });
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void doReadRequest(object usage)
		{
			 Factory.ExecuteMethod(this, "doReadRequest", usage);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void doReadRequest(object usage, object fname)
		{
			 Factory.ExecuteMethod(this, "doReadRequest", usage, fname);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void doReadRequest(object usage, object fname, object domain)
		{
			 Factory.ExecuteMethod(this, "doReadRequest", usage, fname, domain);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		/// <param name="path">optional object path</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void doReadRequest(object usage, object fname, object domain, object path)
		{
			 Factory.ExecuteMethod(this, "doReadRequest", usage, fname, domain, path);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		/// <param name="path">optional object path</param>
		/// <param name="expire">optional object expire</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public void doReadRequest(object usage, object fname, object domain, object path, object expire)
		{
			 Factory.ExecuteMethod(this, "doReadRequest", new object[]{ usage, fname, domain, path, expire });
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public bool doWriteRequest()
		{
			return Factory.ExecuteBoolMethodGet(this, "doWriteRequest");
		}

		#endregion

		#pragma warning restore
	}
}
