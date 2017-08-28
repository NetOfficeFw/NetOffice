using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface Reports 
	/// SupportByVersion MSProject, 11
	/// </summary>
	[SupportByVersion("MSProject", 11)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class Reports : COMObject , IEnumerable<NetOffice.MSProjectApi.Report>
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
                    _type = typeof(Reports);
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Reports(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Reports(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Reports(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Reports(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Reports(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Reports() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Reports(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.MSProjectApi.Report this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Report>(this, "Item", NetOffice.MSProjectApi.Report.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Project Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Project>(this, "Parent", NetOffice.MSProjectApi.Project.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Application>(this, "Application", NetOffice.MSProjectApi.Application.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="newName">string newName</param>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Report Copy(object source, string newName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Report>(this, "Copy", NetOffice.MSProjectApi.Report.LateBindingApiWrapperType, source, newName);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Report Add(string name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Report>(this, "Add", NetOffice.MSProjectApi.Report.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("MSProject", 11)]
		public bool IsPresent(string name)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsPresent", name);
		}

		#endregion

       #region IEnumerable<NetOffice.MSProjectApi.Report> Member
        
        /// <summary>
		/// SupportByVersion MSProject, 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
       public IEnumerator<NetOffice.MSProjectApi.Report> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.MSProjectApi.Report item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersion MSProject, 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}



