using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.VisioApi
{
	///<summary>
	/// Interface LPVISIOWINDOWS 
	/// SupportByVersion Visio, 11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class LPVISIOWINDOWS : COMObject ,IEnumerable<NetOffice.VisioApi.IVWindow>
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
                    _type = typeof(LPVISIOWINDOWS);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIOWINDOWS(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOWINDOWS(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOWINDOWS(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOWINDOWS(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOWINDOWS(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOWINDOWS() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOWINDOWS(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.VisioApi.IVApplication newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVApplication;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ObjectType", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int16 Index</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.VisioApi.IVWindow this[Int16 index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
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
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EventList", paramsArray);
				NetOffice.VisioApi.IVEventList newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVEventList;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 PersistsEvents
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PersistsEvents", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="nID">Int32 nID</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVWindow get_ItemFromID(Int32 nID)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(nID);
			object returnItem = Invoker.PropertyGet(this, "ItemFromID", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ItemFromID
		/// </summary>
		/// <param name="nID">Int32 nID</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow ItemFromID(Int32 nID)
		{
			return get_ItemFromID(nID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="captionOrIndex">object CaptionOrIndex</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVWindow get_ItemEx(object captionOrIndex)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(captionOrIndex);
			object returnItem = Invoker.PropertyGet(this, "ItemEx", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ItemEx
		/// </summary>
		/// <param name="captionOrIndex">object CaptionOrIndex</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow ItemEx(object captionOrIndex)
		{
			return get_ItemEx(captionOrIndex);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void VoidArrange()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "VoidArrange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		/// <param name="nHeight">optional object nHeight</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCaption, nFlags, nType, nLeft, nTop, nWidth, nHeight);
			object returnItem = Invoker.MethodReturn(this, "Add_WithoutMergeArgs", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Add_WithoutMergeArgs", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCaption);
			object returnItem = Invoker.MethodReturn(this, "Add_WithoutMergeArgs", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCaption, nFlags);
			object returnItem = Invoker.MethodReturn(this, "Add_WithoutMergeArgs", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCaption, nFlags, nType);
			object returnItem = Invoker.MethodReturn(this, "Add_WithoutMergeArgs", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType, object nLeft)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCaption, nFlags, nType, nLeft);
			object returnItem = Invoker.MethodReturn(this, "Add_WithoutMergeArgs", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType, object nLeft, object nTop)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCaption, nFlags, nType, nLeft, nTop);
			object returnItem = Invoker.MethodReturn(this, "Add_WithoutMergeArgs", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add_WithoutMergeArgs(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCaption, nFlags, nType, nLeft, nTop, nWidth);
			object returnItem = Invoker.MethodReturn(this, "Add_WithoutMergeArgs", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="nArrangeFlags">optional object nArrangeFlags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Arrange(object nArrangeFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nArrangeFlags);
			Invoker.Method(this, "Arrange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Arrange()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Arrange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		/// <param name="nHeight">optional object nHeight</param>
		/// <param name="bstrMergeID">optional object bstrMergeID</param>
		/// <param name="bstrMergeClass">optional object bstrMergeClass</param>
		/// <param name="nMergePosition">optional object nMergePosition</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight, object bstrMergeID, object bstrMergeClass, object nMergePosition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCaption, nFlags, nType, nLeft, nTop, nWidth, nHeight, bstrMergeID, bstrMergeClass, nMergePosition);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCaption);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCaption, nFlags);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCaption, nFlags, nType);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCaption, nFlags, nType, nLeft);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCaption, nFlags, nType, nLeft, nTop);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCaption, nFlags, nType, nLeft, nTop, nWidth);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		/// <param name="nHeight">optional object nHeight</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCaption, nFlags, nType, nLeft, nTop, nWidth, nHeight);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		/// <param name="nHeight">optional object nHeight</param>
		/// <param name="bstrMergeID">optional object bstrMergeID</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight, object bstrMergeID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCaption, nFlags, nType, nLeft, nTop, nWidth, nHeight, bstrMergeID);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="nFlags">optional object nFlags</param>
		/// <param name="nType">optional object nType</param>
		/// <param name="nLeft">optional object nLeft</param>
		/// <param name="nTop">optional object nTop</param>
		/// <param name="nWidth">optional object nWidth</param>
		/// <param name="nHeight">optional object nHeight</param>
		/// <param name="bstrMergeID">optional object bstrMergeID</param>
		/// <param name="bstrMergeClass">optional object bstrMergeClass</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow Add(object bstrCaption, object nFlags, object nType, object nLeft, object nTop, object nWidth, object nHeight, object bstrMergeID, object bstrMergeClass)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrCaption, nFlags, nType, nLeft, nTop, nWidth, nHeight, bstrMergeID, bstrMergeClass);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.VisioApi.IVWindow> Member
        
        /// <summary>
		/// SupportByVersionAttribute Visio, 11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
       public IEnumerator<NetOffice.VisioApi.IVWindow> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.VisioApi.IVWindow item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Visio, 11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}