using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface Filter 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920593(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Filter : COMObject
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
                    _type = typeof(Filter);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Filter(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Filter(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Filter(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Filter(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Filter(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Filter(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Filter() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Filter(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Application>(this, "Application", NetOffice.MSProjectApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 Index
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Project Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Project>(this, "Parent", NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjItemType FilterType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjItemType>(this, "FilterType");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool ShowInMenu
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowInMenu");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowInMenu", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool ShowRelatedSummaryRows
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowRelatedSummaryRows");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowRelatedSummaryRows", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="highlight">optional bool Highlight = false</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void Apply(object highlight)
		{
			 Factory.ExecuteMethod(this, "Apply", highlight);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public void Apply()
		{
			 Factory.ExecuteMethod(this, "Apply");
		}

		#endregion

		#pragma warning restore
	}
}
