using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.MSProjectApi
{
	///<summary>
	/// DispatchInterface StartDriver 
	/// SupportByVersion MSProject, 11,12,14
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff920699(v=office.14).aspx
	///</summary>
	[SupportByVersionAttribute("MSProject", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class StartDriver : COMObject
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
                    _type = typeof(StartDriver);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public StartDriver(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public StartDriver(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public StartDriver(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public StartDriver(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public StartDriver(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public StartDriver() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public StartDriver(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ActualStartDrivers ActualStartDrivers
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActualStartDrivers", paramsArray);
				NetOffice.MSProjectApi.ActualStartDrivers newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.ActualStartDrivers.LateBindingApiWrapperType) as NetOffice.MSProjectApi.ActualStartDrivers;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.PredecessorDrivers PredecessorDrivers
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PredecessorDrivers", paramsArray);
				NetOffice.MSProjectApi.PredecessorDrivers newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.PredecessorDrivers.LateBindingApiWrapperType) as NetOffice.MSProjectApi.PredecessorDrivers;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ChildDrivers ChildTaskDrivers
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ChildTaskDrivers", paramsArray);
				NetOffice.MSProjectApi.ChildDrivers newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.ChildDrivers.LateBindingApiWrapperType) as NetOffice.MSProjectApi.ChildDrivers;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.CalendarDrivers CalendarDrivers
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CalendarDrivers", paramsArray);
				NetOffice.MSProjectApi.CalendarDrivers newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.CalendarDrivers.LateBindingApiWrapperType) as NetOffice.MSProjectApi.CalendarDrivers;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Task Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				NetOffice.MSProjectApi.Task newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Task.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Task;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.MSProjectApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Application.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,14)]
		public Int32 Suggestions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Suggestions", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,14)]
		public Int32 Warnings
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Warnings", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <param name="overallocationType">NetOffice.MSProjectApi.Enums.PjOverallocationType overallocationType</param>
		[SupportByVersionAttribute("MSProject", 11,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSProjectApi.OverAllocatedAssignments get_OverAllocatedAssignments(NetOffice.MSProjectApi.Enums.PjOverallocationType overallocationType)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(overallocationType);
			object returnItem = Invoker.PropertyGet(this, "OverAllocatedAssignments", paramsArray);
			NetOffice.MSProjectApi.OverAllocatedAssignments newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.OverAllocatedAssignments.LateBindingApiWrapperType) as NetOffice.MSProjectApi.OverAllocatedAssignments;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Alias for get_OverAllocatedAssignments
		/// </summary>
		/// <param name="overallocationType">NetOffice.MSProjectApi.Enums.PjOverallocationType overallocationType</param>
		[SupportByVersionAttribute("MSProject", 11,14)]
		public NetOffice.MSProjectApi.OverAllocatedAssignments OverAllocatedAssignments(NetOffice.MSProjectApi.Enums.PjOverallocationType overallocationType)
		{
			return get_OverAllocatedAssignments(overallocationType);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <param name="startDate">object StartDate</param>
		/// <param name="finishDate">object FinishDate</param>
		[SupportByVersionAttribute("MSProject", 11,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_EffectiveDateDifference(object startDate, object finishDate)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(startDate, finishDate);
			object returnItem = Invoker.PropertyGet(this, "EffectiveDateDifference", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Alias for get_EffectiveDateDifference
		/// </summary>
		/// <param name="startDate">object StartDate</param>
		/// <param name="finishDate">object FinishDate</param>
		[SupportByVersionAttribute("MSProject", 11,14)]
		public object EffectiveDateDifference(object startDate, object finishDate)
		{
			return get_EffectiveDateDifference(startDate, finishDate);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <param name="date">object Date</param>
		/// <param name="duration">object Duration</param>
		[SupportByVersionAttribute("MSProject", 11,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_EffectiveDateAdd(object date, object duration)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(date, duration);
			object returnItem = Invoker.PropertyGet(this, "EffectiveDateAdd", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Alias for get_EffectiveDateAdd
		/// </summary>
		/// <param name="date">object Date</param>
		/// <param name="duration">object Duration</param>
		[SupportByVersionAttribute("MSProject", 11,14)]
		public object EffectiveDateAdd(object date, object duration)
		{
			return get_EffectiveDateAdd(date, duration);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <param name="date">object Date</param>
		/// <param name="duration">object Duration</param>
		[SupportByVersionAttribute("MSProject", 11,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_EffectiveDateSubtract(object date, object duration)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(date, duration);
			object returnItem = Invoker.PropertyGet(this, "EffectiveDateSubtract", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Alias for get_EffectiveDateSubtract
		/// </summary>
		/// <param name="date">object Date</param>
		/// <param name="duration">object Duration</param>
		[SupportByVersionAttribute("MSProject", 11,14)]
		public object EffectiveDateSubtract(object date, object duration)
		{
			return get_EffectiveDateSubtract(date, duration);
		}

		#endregion

		#region Methods

		#endregion
		#pragma warning restore
	}
}