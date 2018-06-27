using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLOpsProfile 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IHTMLOpsProfile : COMObject, NetOffice.MSHTMLApi.IHTMLOpsProfile
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLOpsProfile);
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
                    _type = typeof(IHTMLOpsProfile);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLOpsProfile() : base()
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
		public virtual bool addRequest(string name, object reserved)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "addRequest", name, reserved);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual bool addRequest(string name)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "addRequest", name);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual void clearRequest()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "clearRequest");
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
		public virtual void doRequest(object usage, object fname, object domain, object path, object expire, object reserved)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "doRequest", new object[]{ usage, fname, domain, path, expire, reserved });
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void doRequest(object usage)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "doRequest", usage);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void doRequest(object usage, object fname)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "doRequest", usage, fname);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void doRequest(object usage, object fname, object domain)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "doRequest", usage, fname, domain);
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
		public virtual void doRequest(object usage, object fname, object domain, object path)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "doRequest", usage, fname, domain, path);
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
		public virtual void doRequest(object usage, object fname, object domain, object path, object expire)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "doRequest", new object[]{ usage, fname, domain, path, expire });
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual string getAttribute(string name)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "getAttribute", name);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="value">string value</param>
		/// <param name="prefs">optional object prefs</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool setAttribute(string name, string value, object prefs)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "setAttribute", name, value, prefs);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="value">string value</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual bool setAttribute(string name, string value)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "setAttribute", name, value);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool commitChanges()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "commitChanges");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="reserved">optional object reserved</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool addReadRequest(string name, object reserved)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "addReadRequest", name, reserved);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual bool addReadRequest(string name)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "addReadRequest", name);
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
		public virtual void doReadRequest(object usage, object fname, object domain, object path, object expire, object reserved)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "doReadRequest", new object[]{ usage, fname, domain, path, expire, reserved });
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void doReadRequest(object usage)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "doReadRequest", usage);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void doReadRequest(object usage, object fname)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "doReadRequest", usage, fname);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="usage">object usage</param>
		/// <param name="fname">optional object fname</param>
		/// <param name="domain">optional object domain</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual void doReadRequest(object usage, object fname, object domain)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "doReadRequest", usage, fname, domain);
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
		public virtual void doReadRequest(object usage, object fname, object domain, object path)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "doReadRequest", usage, fname, domain, path);
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
		public virtual void doReadRequest(object usage, object fname, object domain, object path, object expire)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "doReadRequest", new object[]{ usage, fname, domain, path, expire });
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual bool doWriteRequest()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "doWriteRequest");
		}

		#endregion

		#pragma warning restore
	}
}

