using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// Trendlines
	///</summary>
	public class Trendlines_ : COMObject
	{
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines_(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines_(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines_(COMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines_() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		#endregion

		#region Methods

		#endregion

	}

	///<summary>
	/// Interface Trendlines 
	/// SupportByVersion Office, 12,14
	///</summary>
	[SupportByVersionAttribute("Office", 12,14)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class Trendlines : COMObject ,IEnumerable<NetOffice.OfficeApi.IMsoTrendline>
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
                    _type = typeof(Trendlines);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14)]
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
		/// SupportByVersion Office 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 14)]
		public object Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 14)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14
		/// Get
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Office", 14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.OfficeApi.IMsoTrendline this[object index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "_Default", paramsArray);
			NetOffice.OfficeApi.IMsoTrendline newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoTrendline;
			return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object Order</param>
		/// <param name="period">optional object Period</param>
		/// <param name="forward">optional object Forward</param>
		/// <param name="backward">optional object Backward</param>
		/// <param name="intercept">optional object Intercept</param>
		/// <param name="displayEquation">optional object DisplayEquation</param>
		/// <param name="displayRSquared">optional object DisplayRSquared</param>
		/// <param name="name">optional object Name</param>
		[SupportByVersionAttribute("Office", 12,14)]
		public NetOffice.OfficeApi.IMsoTrendline Add(NetOffice.OfficeApi.Enums.XlTrendlineType type, object order, object period, object forward, object backward, object intercept, object displayEquation, object displayRSquared, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, order, period, forward, backward, intercept, displayEquation, displayRSquared, name);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.IMsoTrendline newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoTrendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14)]
		public NetOffice.OfficeApi.IMsoTrendline Add()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.IMsoTrendline newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoTrendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14)]
		public NetOffice.OfficeApi.IMsoTrendline Add(NetOffice.OfficeApi.Enums.XlTrendlineType type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.IMsoTrendline newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoTrendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object Order</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14)]
		public NetOffice.OfficeApi.IMsoTrendline Add(NetOffice.OfficeApi.Enums.XlTrendlineType type, object order)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, order);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.IMsoTrendline newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoTrendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object Order</param>
		/// <param name="period">optional object Period</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14)]
		public NetOffice.OfficeApi.IMsoTrendline Add(NetOffice.OfficeApi.Enums.XlTrendlineType type, object order, object period)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, order, period);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.IMsoTrendline newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoTrendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object Order</param>
		/// <param name="period">optional object Period</param>
		/// <param name="forward">optional object Forward</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14)]
		public NetOffice.OfficeApi.IMsoTrendline Add(NetOffice.OfficeApi.Enums.XlTrendlineType type, object order, object period, object forward)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, order, period, forward);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.IMsoTrendline newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoTrendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object Order</param>
		/// <param name="period">optional object Period</param>
		/// <param name="forward">optional object Forward</param>
		/// <param name="backward">optional object Backward</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14)]
		public NetOffice.OfficeApi.IMsoTrendline Add(NetOffice.OfficeApi.Enums.XlTrendlineType type, object order, object period, object forward, object backward)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, order, period, forward, backward);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.IMsoTrendline newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoTrendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object Order</param>
		/// <param name="period">optional object Period</param>
		/// <param name="forward">optional object Forward</param>
		/// <param name="backward">optional object Backward</param>
		/// <param name="intercept">optional object Intercept</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14)]
		public NetOffice.OfficeApi.IMsoTrendline Add(NetOffice.OfficeApi.Enums.XlTrendlineType type, object order, object period, object forward, object backward, object intercept)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, order, period, forward, backward, intercept);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.IMsoTrendline newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoTrendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object Order</param>
		/// <param name="period">optional object Period</param>
		/// <param name="forward">optional object Forward</param>
		/// <param name="backward">optional object Backward</param>
		/// <param name="intercept">optional object Intercept</param>
		/// <param name="displayEquation">optional object DisplayEquation</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14)]
		public NetOffice.OfficeApi.IMsoTrendline Add(NetOffice.OfficeApi.Enums.XlTrendlineType type, object order, object period, object forward, object backward, object intercept, object displayEquation)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, order, period, forward, backward, intercept, displayEquation);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.IMsoTrendline newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoTrendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object Order</param>
		/// <param name="period">optional object Period</param>
		/// <param name="forward">optional object Forward</param>
		/// <param name="backward">optional object Backward</param>
		/// <param name="intercept">optional object Intercept</param>
		/// <param name="displayEquation">optional object DisplayEquation</param>
		/// <param name="displayRSquared">optional object DisplayRSquared</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14)]
		public NetOffice.OfficeApi.IMsoTrendline Add(NetOffice.OfficeApi.Enums.XlTrendlineType type, object order, object period, object forward, object backward, object intercept, object displayEquation, object displayRSquared)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, order, period, forward, backward, intercept, displayEquation, displayRSquared);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.IMsoTrendline newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType) as NetOffice.OfficeApi.IMsoTrendline;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.OfficeApi.IMsoTrendline> Member
        
        /// <summary>
		/// SupportByVersionAttribute Office, 12,14
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14)]
       public IEnumerator<NetOffice.OfficeApi.IMsoTrendline> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.OfficeApi.IMsoTrendline item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Office, 12,14
		/// </summary>
		[SupportByVersionAttribute("Office", 12,14)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion
		#pragma warning restore
	}
}