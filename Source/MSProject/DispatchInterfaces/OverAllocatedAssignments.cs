using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.MSProjectApi
{
	///<summary>
	/// DispatchInterface OverAllocatedAssignments 
	/// SupportByVersion MSProject, 14
	///</summary>
	[SupportByVersionAttribute("MSProject", 14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class OverAllocatedAssignments : COMObject ,IEnumerable<NetOffice.MSProjectApi.Assignment>
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
                    _type = typeof(OverAllocatedAssignments);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OverAllocatedAssignments(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OverAllocatedAssignments(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OverAllocatedAssignments(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OverAllocatedAssignments() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OverAllocatedAssignments(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 14
		/// Get
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("MSProject", 14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.MSProjectApi.Assignment this[Int32 index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.MSProjectApi.Assignment newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Assignment.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Assignment;
			return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14)]
		public Int32 TotalDetectedCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TotalDetectedCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14)]
		public NetOffice.MSProjectApi.Task Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				NetOffice.MSProjectApi.Task newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Task.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Task;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14)]
		public NetOffice.MSProjectApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.MSProjectApi.Application newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Application.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Application;
				return newObject;
			}
		}

		#endregion

		#region Methods

		#endregion

       #region IEnumerable<NetOffice.MSProjectApi.Assignment> Member
        
        /// <summary>
		/// SupportByVersionAttribute MSProject, 14
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14)]
       public IEnumerator<NetOffice.MSProjectApi.Assignment> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.MSProjectApi.Assignment item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute MSProject, 14
		/// </summary>
		[SupportByVersionAttribute("MSProject", 14)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}