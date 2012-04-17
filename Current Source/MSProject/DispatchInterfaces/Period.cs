using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using LateBindingApi.Core;
namespace NetOffice.MSProjectApi
{
	///<summary>
	/// DispatchInterface Period 
	/// SupportByVersion MSProject, 12,14
	///</summary>
	[SupportByVersionAttribute("MSProject", 12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Period : COMObject
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
                    _type = typeof(Period);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Period(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Period(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Period(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Period() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Period(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 12,14)]
		public NetOffice.MSProjectApi.Calendar Calendar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Calendar", paramsArray);
				NetOffice.MSProjectApi.Calendar newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Calendar.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Calendar;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 12,14)]
		public NetOffice.MSProjectApi.Shift Shift1
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shift1", paramsArray);
				NetOffice.MSProjectApi.Shift newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Shift.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shift;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 12,14)]
		public NetOffice.MSProjectApi.Shift Shift2
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shift2", paramsArray);
				NetOffice.MSProjectApi.Shift newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Shift.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shift;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 12,14)]
		public NetOffice.MSProjectApi.Shift Shift3
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shift3", paramsArray);
				NetOffice.MSProjectApi.Shift newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Shift.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shift;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 12,14)]
		public bool Working
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Working", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Working", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 12,14)]
		public Int16 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 12,14)]
		public NetOffice.MSProjectApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.MSProjectApi.Application newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Application.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 12,14)]
		public NetOffice.MSProjectApi.Calendar Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				NetOffice.MSProjectApi.Calendar newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Calendar.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Calendar;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 12,14)]
		public NetOffice.MSProjectApi.Shift Shift4
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shift4", paramsArray);
				NetOffice.MSProjectApi.Shift newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Shift.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shift;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 12,14)]
		public NetOffice.MSProjectApi.Shift Shift5
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shift5", paramsArray);
				NetOffice.MSProjectApi.Shift newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Shift.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shift;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 12, 14
		/// </summary>
		[SupportByVersionAttribute("MSProject", 12,14)]
		public void Default()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Default", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}