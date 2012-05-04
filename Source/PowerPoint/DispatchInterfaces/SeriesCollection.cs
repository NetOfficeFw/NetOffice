using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.PowerPointApi
{
	///<summary>
	/// DispatchInterface SeriesCollection 
	/// SupportByVersion PowerPoint, 14
	///</summary>
	[SupportByVersionAttribute("PowerPoint", 14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class SeriesCollection : COMObject ,IEnumerable<NetOffice.PowerPointApi.Series>
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
                    _type = typeof(SeriesCollection);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SeriesCollection(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SeriesCollection(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SeriesCollection(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SeriesCollection() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SeriesCollection(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
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
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
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
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
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
		/// SupportByVersion PowerPoint 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.PowerPointApi.Application newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Application.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Application;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="source">object Source</param>
		/// <param name="rowcol">optional object Rowcol</param>
		/// <param name="categoryLabels">optional object CategoryLabels</param>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public object Extend(object source, object rowcol, object categoryLabels)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, rowcol, categoryLabels);
			object returnItem = Invoker.MethodReturn(this, "Extend", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="source">object Source</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public object Extend(object source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source);
			object returnItem = Invoker.MethodReturn(this, "Extend", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="source">object Source</param>
		/// <param name="rowcol">optional object Rowcol</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public object Extend(object source, object rowcol)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, rowcol);
			object returnItem = Invoker.MethodReturn(this, "Extend", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.Series NewSeries()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "NewSeries", paramsArray);
			NetOffice.PowerPointApi.Series newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Series.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Series;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="source">object Source</param>
		/// <param name="rowcol">optional NetOffice.PowerPointApi.Enums.XlRowCol Rowcol = -4105</param>
		/// <param name="seriesLabels">optional object SeriesLabels</param>
		/// <param name="categoryLabels">optional object CategoryLabels</param>
		/// <param name="replace">optional object Replace</param>
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.Series Add(object source, object rowcol, object seriesLabels, object categoryLabels, object replace)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, rowcol, seriesLabels, categoryLabels, replace);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.PowerPointApi.Series newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Series.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Series;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="source">object Source</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.Series Add(object source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.PowerPointApi.Series newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Series.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Series;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="source">object Source</param>
		/// <param name="rowcol">optional NetOffice.PowerPointApi.Enums.XlRowCol Rowcol = -4105</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.Series Add(object source, object rowcol)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, rowcol);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.PowerPointApi.Series newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Series.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Series;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="source">object Source</param>
		/// <param name="rowcol">optional NetOffice.PowerPointApi.Enums.XlRowCol Rowcol = -4105</param>
		/// <param name="seriesLabels">optional object SeriesLabels</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.Series Add(object source, object rowcol, object seriesLabels)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, rowcol, seriesLabels);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.PowerPointApi.Series newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Series.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Series;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="source">object Source</param>
		/// <param name="rowcol">optional NetOffice.PowerPointApi.Enums.XlRowCol Rowcol = -4105</param>
		/// <param name="seriesLabels">optional object SeriesLabels</param>
		/// <param name="categoryLabels">optional object CategoryLabels</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 14)]
		public NetOffice.PowerPointApi.Series Add(object source, object rowcol, object seriesLabels, object categoryLabels)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, rowcol, seriesLabels, categoryLabels);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.PowerPointApi.Series newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Series.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Series;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("PowerPoint", 14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.PowerPointApi.Series this[object index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "_Default", paramsArray);
				NetOffice.PowerPointApi.Series newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.Series.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Series;
				return newObject;
			}
		}

		#endregion

       #region IEnumerable<NetOffice.PowerPointApi.Series> Member
        
        /// <summary>
		/// SupportByVersionAttribute PowerPoint, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
       public IEnumerator<NetOffice.PowerPointApi.Series> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.PowerPointApi.Series item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute PowerPoint, 14
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 14)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion
		#pragma warning restore
	}
}