using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface Hyperlink 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Hyperlink : COMObject
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
                    _type = typeof(Hyperlink);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Hyperlink(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Hyperlink(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlink(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlink(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlink(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlink(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlink() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlink(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api.ISpreadsheet Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.ISpreadsheet>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string Address
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Address");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Address", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public NetOffice.OWC10Api._Range Parent
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api._Range>(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string SubAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SubAddress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SubAddress", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		/// <param name="method">optional object method</param>
		/// <param name="headerInfo">optional object headerInfo</param>
		[SupportByVersion("OWC10", 1)]
		public void Follow(object newWindow, object addHistory, object extraInfo, object method, object headerInfo)
		{
			 Factory.ExecuteMethod(this, "Follow", new object[]{ newWindow, addHistory, extraInfo, method, headerInfo });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Follow()
		{
			 Factory.ExecuteMethod(this, "Follow");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="newWindow">optional object newWindow</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Follow(object newWindow)
		{
			 Factory.ExecuteMethod(this, "Follow", newWindow);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Follow(object newWindow, object addHistory)
		{
			 Factory.ExecuteMethod(this, "Follow", newWindow, addHistory);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Follow(object newWindow, object addHistory, object extraInfo)
		{
			 Factory.ExecuteMethod(this, "Follow", newWindow, addHistory, extraInfo);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="newWindow">optional object newWindow</param>
		/// <param name="addHistory">optional object addHistory</param>
		/// <param name="extraInfo">optional object extraInfo</param>
		/// <param name="method">optional object method</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Follow(object newWindow, object addHistory, object extraInfo, object method)
		{
			 Factory.ExecuteMethod(this, "Follow", newWindow, addHistory, extraInfo, method);
		}

		#endregion

		#pragma warning restore
	}
}
