using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface Trendlines 
	/// SupportByVersion Word, 14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845131.aspx
	///</summary>
	[SupportByVersionAttribute("Word", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Trendlines : COMObject ,IEnumerable<NetOffice.WordApi.Trendline>
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

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Trendlines(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines(ICOMObject replacedObject) : base(replacedObject)
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823212.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821991.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822618.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public object Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192418.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820757.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object Order</param>
		/// <param name="period">optional object Period</param>
		/// <param name="forward">optional object Forward</param>
		/// <param name="backward">optional object Backward</param>
		/// <param name="intercept">optional object Intercept</param>
		/// <param name="displayEquation">optional object DisplayEquation</param>
		/// <param name="displayRSquared">optional object DisplayRSquared</param>
		/// <param name="name">optional object Name</param>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation, object displayRSquared, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, order, period, forward, backward, intercept, displayEquation, displayRSquared, name);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Trendline newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Trendline.LateBindingApiWrapperType) as NetOffice.WordApi.Trendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820757.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Trendline newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Trendline.LateBindingApiWrapperType) as NetOffice.WordApi.Trendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820757.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlTrendlineType Type = -4132</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Trendline newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Trendline.LateBindingApiWrapperType) as NetOffice.WordApi.Trendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820757.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object Order</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add(object type, object order)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, order);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Trendline newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Trendline.LateBindingApiWrapperType) as NetOffice.WordApi.Trendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820757.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object Order</param>
		/// <param name="period">optional object Period</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add(object type, object order, object period)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, order, period);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Trendline newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Trendline.LateBindingApiWrapperType) as NetOffice.WordApi.Trendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820757.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object Order</param>
		/// <param name="period">optional object Period</param>
		/// <param name="forward">optional object Forward</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add(object type, object order, object period, object forward)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, order, period, forward);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Trendline newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Trendline.LateBindingApiWrapperType) as NetOffice.WordApi.Trendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820757.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object Order</param>
		/// <param name="period">optional object Period</param>
		/// <param name="forward">optional object Forward</param>
		/// <param name="backward">optional object Backward</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add(object type, object order, object period, object forward, object backward)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, order, period, forward, backward);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Trendline newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Trendline.LateBindingApiWrapperType) as NetOffice.WordApi.Trendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820757.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object Order</param>
		/// <param name="period">optional object Period</param>
		/// <param name="forward">optional object Forward</param>
		/// <param name="backward">optional object Backward</param>
		/// <param name="intercept">optional object Intercept</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add(object type, object order, object period, object forward, object backward, object intercept)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, order, period, forward, backward, intercept);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Trendline newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Trendline.LateBindingApiWrapperType) as NetOffice.WordApi.Trendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820757.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object Order</param>
		/// <param name="period">optional object Period</param>
		/// <param name="forward">optional object Forward</param>
		/// <param name="backward">optional object Backward</param>
		/// <param name="intercept">optional object Intercept</param>
		/// <param name="displayEquation">optional object DisplayEquation</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, order, period, forward, backward, intercept, displayEquation);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Trendline newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Trendline.LateBindingApiWrapperType) as NetOffice.WordApi.Trendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820757.aspx
		/// </summary>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object Order</param>
		/// <param name="period">optional object Period</param>
		/// <param name="forward">optional object Forward</param>
		/// <param name="backward">optional object Backward</param>
		/// <param name="intercept">optional object Intercept</param>
		/// <param name="displayEquation">optional object DisplayEquation</param>
		/// <param name="displayRSquared">optional object DisplayRSquared</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation, object displayRSquared)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, order, period, forward, backward, intercept, displayEquation, displayRSquared);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Trendline newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Trendline.LateBindingApiWrapperType) as NetOffice.WordApi.Trendline;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.WordApi.Trendline this[object index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "_Default", paramsArray);
				NetOffice.WordApi.Trendline newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Trendline.LateBindingApiWrapperType) as NetOffice.WordApi.Trendline;
				return newObject;
			}
		}

		#endregion

       #region IEnumerable<NetOffice.WordApi.Trendline> Member
        
        /// <summary>
		/// SupportByVersionAttribute Word, 14,15,16
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
       public IEnumerator<NetOffice.WordApi.Trendline> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.WordApi.Trendline item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Word, 14,15,16
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion
		#pragma warning restore
	}
}