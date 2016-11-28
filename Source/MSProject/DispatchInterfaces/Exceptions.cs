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
	/// DispatchInterface Exceptions 
	/// SupportByVersion MSProject, 11,12,14
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff920590(v=office.14).aspx
	///</summary>
	[SupportByVersionAttribute("MSProject", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Exceptions : COMObject ,IEnumerable<NetOffice.MSProjectApi.Exception>
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
                    _type = typeof(Exceptions);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Exceptions(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Exceptions(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Exceptions(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Exceptions(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Exceptions(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Exceptions() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Exceptions(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

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
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Calendar Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				NetOffice.MSProjectApi.Calendar newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Calendar.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Calendar;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
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
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.MSProjectApi.Exception this[object index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.MSProjectApi.Exception newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Exception;
			return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType Type</param>
		/// <param name="start">object Start</param>
		/// <param name="finish">optional object Finish</param>
		/// <param name="occurrences">optional object Occurrences</param>
		/// <param name="name">optional object Name</param>
		/// <param name="period">optional object Period</param>
		/// <param name="daysOfWeek">optional object DaysOfWeek</param>
		/// <param name="monthPosition">optional object MonthPosition</param>
		/// <param name="monthItem">optional object MonthItem</param>
		/// <param name="month">optional object Month</param>
		/// <param name="monthDay">optional object MonthDay</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek, object monthPosition, object monthItem, object month, object monthDay)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, start, finish, occurrences, name, period, daysOfWeek, monthPosition, monthItem, month, monthDay);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.Exception newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Exception;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType Type</param>
		/// <param name="start">object Start</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, start);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.Exception newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Exception;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType Type</param>
		/// <param name="start">object Start</param>
		/// <param name="finish">optional object Finish</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, start, finish);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.Exception newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Exception;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType Type</param>
		/// <param name="start">object Start</param>
		/// <param name="finish">optional object Finish</param>
		/// <param name="occurrences">optional object Occurrences</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, start, finish, occurrences);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.Exception newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Exception;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType Type</param>
		/// <param name="start">object Start</param>
		/// <param name="finish">optional object Finish</param>
		/// <param name="occurrences">optional object Occurrences</param>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, start, finish, occurrences, name);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.Exception newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Exception;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType Type</param>
		/// <param name="start">object Start</param>
		/// <param name="finish">optional object Finish</param>
		/// <param name="occurrences">optional object Occurrences</param>
		/// <param name="name">optional object Name</param>
		/// <param name="period">optional object Period</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, start, finish, occurrences, name, period);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.Exception newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Exception;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType Type</param>
		/// <param name="start">object Start</param>
		/// <param name="finish">optional object Finish</param>
		/// <param name="occurrences">optional object Occurrences</param>
		/// <param name="name">optional object Name</param>
		/// <param name="period">optional object Period</param>
		/// <param name="daysOfWeek">optional object DaysOfWeek</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, start, finish, occurrences, name, period, daysOfWeek);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.Exception newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Exception;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType Type</param>
		/// <param name="start">object Start</param>
		/// <param name="finish">optional object Finish</param>
		/// <param name="occurrences">optional object Occurrences</param>
		/// <param name="name">optional object Name</param>
		/// <param name="period">optional object Period</param>
		/// <param name="daysOfWeek">optional object DaysOfWeek</param>
		/// <param name="monthPosition">optional object MonthPosition</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek, object monthPosition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, start, finish, occurrences, name, period, daysOfWeek, monthPosition);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.Exception newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Exception;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType Type</param>
		/// <param name="start">object Start</param>
		/// <param name="finish">optional object Finish</param>
		/// <param name="occurrences">optional object Occurrences</param>
		/// <param name="name">optional object Name</param>
		/// <param name="period">optional object Period</param>
		/// <param name="daysOfWeek">optional object DaysOfWeek</param>
		/// <param name="monthPosition">optional object MonthPosition</param>
		/// <param name="monthItem">optional object MonthItem</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek, object monthPosition, object monthItem)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, start, finish, occurrences, name, period, daysOfWeek, monthPosition, monthItem);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.Exception newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Exception;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="type">NetOffice.MSProjectApi.Enums.PjExceptionType Type</param>
		/// <param name="start">object Start</param>
		/// <param name="finish">optional object Finish</param>
		/// <param name="occurrences">optional object Occurrences</param>
		/// <param name="name">optional object Name</param>
		/// <param name="period">optional object Period</param>
		/// <param name="daysOfWeek">optional object DaysOfWeek</param>
		/// <param name="monthPosition">optional object MonthPosition</param>
		/// <param name="monthItem">optional object MonthItem</param>
		/// <param name="month">optional object Month</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Exception Add(NetOffice.MSProjectApi.Enums.PjExceptionType type, object start, object finish, object occurrences, object name, object period, object daysOfWeek, object monthPosition, object monthItem, object month)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, start, finish, occurrences, name, period, daysOfWeek, monthPosition, monthItem, month);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.Exception newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Exception.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Exception;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.MSProjectApi.Exception> Member
        
        /// <summary>
		/// SupportByVersionAttribute MSProject, 11,12,14
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
       public IEnumerator<NetOffice.MSProjectApi.Exception> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.MSProjectApi.Exception item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute MSProject, 11,12,14
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}