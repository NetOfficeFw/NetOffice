using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface Group2 
	/// SupportByVersion MSProject, 11,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920602(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,14)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Group2 : COMObject
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
                    _type = typeof(Group2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Group2(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Group2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Group2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Group2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Group2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Group2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Group2() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Group2(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public NetOffice.MSProjectApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Application>(this, "Application", NetOffice.MSProjectApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public Int32 Index
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public NetOffice.MSProjectApi.Project Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Project>(this, "Parent", NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public NetOffice.MSProjectApi.GroupCriteria2 GroupCriteria
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.GroupCriteria2>(this, "GroupCriteria", NetOffice.MSProjectApi.GroupCriteria2.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "GroupCriteria", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public bool ShowSummary
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowSummary");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowSummary", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public bool GroupAssignments
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "GroupAssignments");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GroupAssignments", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public bool MaintainHierarchy
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MaintainHierarchy");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaintainHierarchy", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		#endregion

		#pragma warning restore
	}
}
