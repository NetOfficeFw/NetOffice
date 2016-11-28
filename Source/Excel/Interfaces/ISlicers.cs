using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.ExcelApi
{
	///<summary>
	/// Interface ISlicers 
	/// SupportByVersion Excel, 14,15,16
	///</summary>
	[SupportByVersionAttribute("Excel", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class ISlicers : COMObject ,IEnumerable<NetOffice.ExcelApi.Slicer>
	{
		#pragma warning disable
		#region Type Information

        /// <summary>
        /// Instance Type
        /// </summary>
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
                    _type = typeof(ISlicers);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ISlicers(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISlicers(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISlicers(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISlicers(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISlicers(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISlicers() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISlicers(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.ExcelApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Application.LateBindingApiWrapperType) as NetOffice.ExcelApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlCreator)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
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
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
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
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.ExcelApi.Slicer this[object index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "_Default", paramsArray);
			NetOffice.ExcelApi.Slicer newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType) as NetOffice.ExcelApi.Slicer;
			return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="slicerDestination">object SlicerDestination</param>
		/// <param name="level">optional object Level</param>
		/// <param name="name">optional object Name</param>
		/// <param name="caption">optional object Caption</param>
		/// <param name="top">optional object Top</param>
		/// <param name="left">optional object Left</param>
		/// <param name="width">optional object Width</param>
		/// <param name="height">optional object Height</param>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption, object top, object left, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(slicerDestination, level, name, caption, top, left, width, height);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.Slicer newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType) as NetOffice.ExcelApi.Slicer;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="slicerDestination">object SlicerDestination</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicer Add(object slicerDestination)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(slicerDestination);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.Slicer newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType) as NetOffice.ExcelApi.Slicer;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="slicerDestination">object SlicerDestination</param>
		/// <param name="level">optional object Level</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(slicerDestination, level);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.Slicer newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType) as NetOffice.ExcelApi.Slicer;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="slicerDestination">object SlicerDestination</param>
		/// <param name="level">optional object Level</param>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(slicerDestination, level, name);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.Slicer newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType) as NetOffice.ExcelApi.Slicer;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="slicerDestination">object SlicerDestination</param>
		/// <param name="level">optional object Level</param>
		/// <param name="name">optional object Name</param>
		/// <param name="caption">optional object Caption</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(slicerDestination, level, name, caption);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.Slicer newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType) as NetOffice.ExcelApi.Slicer;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="slicerDestination">object SlicerDestination</param>
		/// <param name="level">optional object Level</param>
		/// <param name="name">optional object Name</param>
		/// <param name="caption">optional object Caption</param>
		/// <param name="top">optional object Top</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(slicerDestination, level, name, caption, top);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.Slicer newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType) as NetOffice.ExcelApi.Slicer;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="slicerDestination">object SlicerDestination</param>
		/// <param name="level">optional object Level</param>
		/// <param name="name">optional object Name</param>
		/// <param name="caption">optional object Caption</param>
		/// <param name="top">optional object Top</param>
		/// <param name="left">optional object Left</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption, object top, object left)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(slicerDestination, level, name, caption, top, left);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.Slicer newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType) as NetOffice.ExcelApi.Slicer;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// 
		/// </summary>
		/// <param name="slicerDestination">object SlicerDestination</param>
		/// <param name="level">optional object Level</param>
		/// <param name="name">optional object Name</param>
		/// <param name="caption">optional object Caption</param>
		/// <param name="top">optional object Top</param>
		/// <param name="left">optional object Left</param>
		/// <param name="width">optional object Width</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Slicer Add(object slicerDestination, object level, object name, object caption, object top, object left, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(slicerDestination, level, name, caption, top, left, width);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.Slicer newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Slicer.LateBindingApiWrapperType) as NetOffice.ExcelApi.Slicer;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.ExcelApi.Slicer> Member
        
        /// <summary>
		/// SupportByVersionAttribute Excel, 14,15,16
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
       public IEnumerator<NetOffice.ExcelApi.Slicer> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.ExcelApi.Slicer item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Excel, 14,15,16
		/// </summary>
		[SupportByVersionAttribute("Excel", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}