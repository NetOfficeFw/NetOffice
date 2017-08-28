using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface LPVISIOEVENT 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class LPVISIOEVENT : COMObject
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
                    _type = typeof(LPVISIOEVENT);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public LPVISIOEVENT(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIOEVENT(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOEVENT(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOEVENT(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOEVENT(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOEVENT(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOEVENT() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOEVENT(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVEventList>(this, "EventList");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Index
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Event
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Event");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Event", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Action
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Action");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Action", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string Target
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Target");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Target", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string TargetArgs
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TargetArgs");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TargetArgs", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 ID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Enabled
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Enabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Enabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Persistable
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Persistable");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Persistent
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Persistent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Persistent", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="contextString">string contextString</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Trigger(string contextString)
		{
			 Factory.ExecuteMethod(this, "Trigger", contextString);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32[] GetFilterObjects()
		{
			object[] paramsArray = null;
			object returnItem = (object)Invoker.MethodReturn(this, "GetFilterObjects", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objects">Int32[] objects</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SetFilterObjects(Int32[] objects)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)objects);
            Invoker.Method(this, "SetFilterObjects", paramsArray);
        }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32[] GetFilterCommands()
		{
			object[] paramsArray = null;
			object returnItem = (object)Invoker.MethodReturn(this, "GetFilterCommands", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="commands">Int32[] commands</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SetFilterCommands(Int32[] commands)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)commands);
            Invoker.Method(this, "SetFilterCommands", paramsArray);
        }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16[] GetFilterSRC()
		{
			object[] paramsArray = null;
			object returnItem = (object)Invoker.MethodReturn(this, "GetFilterSRC", paramsArray);
			return (Int16[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sRCStream">Int16[] sRCStream</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SetFilterSRC(Int16[] sRCStream)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)sRCStream);
            Invoker.Method(this, "SetFilterSRC", paramsArray);
        }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public Int32[] GetFilterActions()
		{
			object[] paramsArray = null;
			object returnItem = (object)Invoker.MethodReturn(this, "GetFilterActions", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="filterActionStream">Int32[] filterActionStream</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public void SetFilterActions(Int32[] filterActionStream)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)filterActionStream);
            Invoker.Method(this, "SetFilterActions", paramsArray);
        }

		#endregion

		#pragma warning restore
	}
}
