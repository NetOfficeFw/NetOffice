using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.VisioApi;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// Interface LPVISIOEVENT 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class LPVISIOEVENT : COMObject, NetOffice.VisioApi.LPVISIOEVENT
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
                    _contractType = typeof(NetOffice.VisioApi.LPVISIOEVENT);
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
                    _type = typeof(LPVISIOEVENT);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public LPVISIOEVENT() : base()
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
		public virtual NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVEventList>(this, "EventList");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 Index
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 Event
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Event");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Event", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 Action
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Action");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Action", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string Target
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Target");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Target", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string TargetArgs
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "TargetArgs");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TargetArgs", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 ObjectType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 ID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 Enabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Enabled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Enabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 Persistable
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Persistable");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 Persistent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Persistent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Persistent", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="contextString">string contextString</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Trigger(string contextString)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Trigger", contextString);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Delete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32[] GetFilterObjects()
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
		public virtual void SetFilterObjects(Int32[] objects)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)objects);
            Invoker.Method(this, "SetFilterObjects", paramsArray);
        }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32[] GetFilterCommands()
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
		public virtual void SetFilterCommands(Int32[] commands)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)commands);
            Invoker.Method(this, "SetFilterCommands", paramsArray);
        }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16[] GetFilterSRC()
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
		public virtual void SetFilterSRC(Int16[] sRCStream)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)sRCStream);
            Invoker.Method(this, "SetFilterSRC", paramsArray);
        }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual Int32[] GetFilterActions()
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
		public virtual void SetFilterActions(Int32[] filterActionStream)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)filterActionStream);
            Invoker.Method(this, "SetFilterActions", paramsArray);
        }

		#endregion

		#pragma warning restore
	}
}

