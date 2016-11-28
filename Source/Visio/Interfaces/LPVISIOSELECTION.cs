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
	/// Interface LPVISIOSELECTION 
	/// SupportByVersion Visio, 11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class LPVISIOSELECTION : COMObject ,IEnumerable<NetOffice.VisioApi.IVShape>
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
                    _type = typeof(LPVISIOSELECTION);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIOSELECTION(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSELECTION(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSELECTION(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSELECTION(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSELECTION(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSELECTION() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSELECTION(string progId) : base(progId)
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
		public Int16 Stat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Stat", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVShape get_Item16(Int16 index)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item16", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Item16
		/// </summary>
		/// <param name="index">Int16 Index</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape Item16(Int16 index)
		{
			return get_Item16(index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 Count16
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count16", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Document", paramsArray);
				NetOffice.VisioApi.IVDocument newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVDocument;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVPage ContainingPage
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContainingPage", paramsArray);
				NetOffice.VisioApi.IVPage newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVPage;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVMaster ContainingMaster
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContainingMaster", paramsArray);
				NetOffice.VisioApi.IVMaster newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVMaster;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape ContainingShape
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContainingShape", paramsArray);
				NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Style
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Style", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Style", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string StyleKeepFmt
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StyleKeepFmt", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "StyleKeepFmt", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string LineStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LineStyle", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LineStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string LineStyleKeepFmt
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LineStyleKeepFmt", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LineStyleKeepFmt", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string FillStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FillStyle", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FillStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string FillStyleKeepFmt
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FillStyleKeepFmt", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FillStyleKeepFmt", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string TextStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TextStyle", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TextStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string TextStyleKeepFmt
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TextStyleKeepFmt", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TextStyleKeepFmt", paramsArray);
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
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.VisioApi.IVShape this[Int32 index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
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
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 IterationMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IterationMode", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "IterationMode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_ItemStatus(Int32 index)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "ItemStatus", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ItemStatus
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 ItemStatus(Int32 index)
		{
			return get_ItemStatus(index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape PrimaryItem
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrimaryItem", paramsArray);
				NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public stdole.Picture Picture
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Picture", paramsArray);
				stdole.Picture newObject = Factory.CreateObjectFromComProxy(this,returnItem) as stdole.Picture;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 ContainingPageID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContainingPageID", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 ContainingMasterID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContainingMasterID", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.IVMaster DataGraphic
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DataGraphic", paramsArray);
				NetOffice.VisioApi.IVMaster newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVMaster;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DataGraphic", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVSelection SelectionForDragCopy
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SelectionForDragCopy", paramsArray);
				NetOffice.VisioApi.IVSelection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVSelection;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Export(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "Export", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void BringForward()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "BringForward", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void BringToFront()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "BringToFront", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void SendBackward()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendBackward", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void SendToBack()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendToBack", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Combine()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Combine", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Fragment()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Fragment", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Intersect()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Intersect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Subtract()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Subtract", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Union()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Union", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void FlipHorizontal()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "FlipHorizontal", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void FlipVertical()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "FlipVertical", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void ReverseEnds()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ReverseEnds", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Rotate90()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Rotate90", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void old_Copy()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "old_Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void old_Cut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "old_Cut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void VoidDuplicate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "VoidDuplicate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void VoidGroup()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "VoidGroup", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void ConvertToGroup()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ConvertToGroup", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Ungroup()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Ungroup", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void SelectAll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SelectAll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void DeselectAll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DeselectAll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sheetObject">NetOffice.VisioApi.IVShape SheetObject</param>
		/// <param name="selectAction">Int16 SelectAction</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Select(NetOffice.VisioApi.IVShape sheetObject, Int16 selectAction)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sheetObject, selectAction);
			Invoker.Method(this, "Select", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Trim()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Trim", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Join()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Join", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="tolerance">Double Tolerance</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void FitCurve(Double tolerance, Int16 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tolerance, flags);
			Invoker.Method(this, "FitCurve", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Layout()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Layout", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="lpr8Left">Double lpr8Left</param>
		/// <param name="lpr8Bottom">Double lpr8Bottom</param>
		/// <param name="lpr8Right">Double lpr8Right</param>
		/// <param name="lpr8Top">Double lpr8Top</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void BoundingBox(Int16 flags, out Double lpr8Left, out Double lpr8Bottom, out Double lpr8Right, out Double lpr8Top)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true,true,true);
			lpr8Left = 0;
			lpr8Bottom = 0;
			lpr8Right = 0;
			lpr8Top = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(flags, lpr8Left, lpr8Bottom, lpr8Right, lpr8Top);
			Invoker.Method(this, "BoundingBox", paramsArray, modifiers);
			lpr8Left = (Double)paramsArray[1];
			lpr8Bottom = (Double)paramsArray[2];
			lpr8Right = (Double)paramsArray[3];
			lpr8Top = (Double)paramsArray[4];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="tolerance">Double Tolerance</param>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="x">optional object x</param>
		/// <param name="y">optional object y</param>
		/// <param name="resultsMaster">optional object ResultsMaster</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawRegion(Double tolerance, Int16 flags, object x, object y, object resultsMaster)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tolerance, flags, x, y, resultsMaster);
			object returnItem = Invoker.MethodReturn(this, "DrawRegion", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="tolerance">Double Tolerance</param>
		/// <param name="flags">Int16 Flags</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawRegion(Double tolerance, Int16 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tolerance, flags);
			object returnItem = Invoker.MethodReturn(this, "DrawRegion", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="tolerance">Double Tolerance</param>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="x">optional object x</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawRegion(Double tolerance, Int16 flags, object x)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tolerance, flags, x);
			object returnItem = Invoker.MethodReturn(this, "DrawRegion", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="tolerance">Double Tolerance</param>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="x">optional object x</param>
		/// <param name="y">optional object y</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawRegion(Double tolerance, Int16 flags, object x, object y)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tolerance, flags, x, y);
			object returnItem = Invoker.MethodReturn(this, "DrawRegion", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape Group()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Group", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void SwapEnds()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SwapEnds", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void AddToGroup()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AddToGroup", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void RemoveFromGroup()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RemoveFromGroup", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVSelection Duplicate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Duplicate", paramsArray);
			NetOffice.VisioApi.IVSelection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVSelection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="flags">optional object Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Copy(object flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(flags);
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Copy()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="flags">optional object Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Cut(object flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(flags);
			Invoker.Method(this, "Cut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Cut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Cut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dx">Double dx</param>
		/// <param name="dy">Double dy</param>
		/// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Move(Double dx, Double dy, object unitsNameOrCode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dx, dy, unitsNameOrCode);
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dx">Double dx</param>
		/// <param name="dy">Double dy</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Move(Double dx, Double dy)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dx, dy);
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="angle">Double Angle</param>
		/// <param name="angleUnitsNameOrCode">optional object AngleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="rotationType">optional NetOffice.VisioApi.Enums.VisRotationTypes RotationType = 0</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		/// <param name="pinY">optional Double PinY = 0</param>
		/// <param name="pinUnitsNameOrCode">optional object PinUnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards, object rotationType, object pinX, object pinY, object pinUnitsNameOrCode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(angle, angleUnitsNameOrCode, blastGuards, rotationType, pinX, pinY, pinUnitsNameOrCode);
			Invoker.Method(this, "Rotate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="angle">Double Angle</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Rotate(Double angle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(angle);
			Invoker.Method(this, "Rotate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="angle">Double Angle</param>
		/// <param name="angleUnitsNameOrCode">optional object AngleUnitsNameOrCode</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Rotate(Double angle, object angleUnitsNameOrCode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(angle, angleUnitsNameOrCode);
			Invoker.Method(this, "Rotate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="angle">Double Angle</param>
		/// <param name="angleUnitsNameOrCode">optional object AngleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(angle, angleUnitsNameOrCode, blastGuards);
			Invoker.Method(this, "Rotate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="angle">Double Angle</param>
		/// <param name="angleUnitsNameOrCode">optional object AngleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="rotationType">optional NetOffice.VisioApi.Enums.VisRotationTypes RotationType = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards, object rotationType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(angle, angleUnitsNameOrCode, blastGuards, rotationType);
			Invoker.Method(this, "Rotate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="angle">Double Angle</param>
		/// <param name="angleUnitsNameOrCode">optional object AngleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="rotationType">optional NetOffice.VisioApi.Enums.VisRotationTypes RotationType = 0</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards, object rotationType, object pinX)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(angle, angleUnitsNameOrCode, blastGuards, rotationType, pinX);
			Invoker.Method(this, "Rotate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="angle">Double Angle</param>
		/// <param name="angleUnitsNameOrCode">optional object AngleUnitsNameOrCode</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="rotationType">optional NetOffice.VisioApi.Enums.VisRotationTypes RotationType = 0</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		/// <param name="pinY">optional Double PinY = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Rotate(Double angle, object angleUnitsNameOrCode, object blastGuards, object rotationType, object pinX, object pinY)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(angle, angleUnitsNameOrCode, blastGuards, rotationType, pinX, pinY);
			Invoker.Method(this, "Rotate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="alignHorizontal">NetOffice.VisioApi.Enums.VisHorizontalAlignTypes AlignHorizontal</param>
		/// <param name="alignVertical">NetOffice.VisioApi.Enums.VisVerticalAlignTypes AlignVertical</param>
		/// <param name="glueToGuide">optional bool GlueToGuide = false</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Align(NetOffice.VisioApi.Enums.VisHorizontalAlignTypes alignHorizontal, NetOffice.VisioApi.Enums.VisVerticalAlignTypes alignVertical, object glueToGuide)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(alignHorizontal, alignVertical, glueToGuide);
			Invoker.Method(this, "Align", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="alignHorizontal">NetOffice.VisioApi.Enums.VisHorizontalAlignTypes AlignHorizontal</param>
		/// <param name="alignVertical">NetOffice.VisioApi.Enums.VisVerticalAlignTypes AlignVertical</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Align(NetOffice.VisioApi.Enums.VisHorizontalAlignTypes alignHorizontal, NetOffice.VisioApi.Enums.VisVerticalAlignTypes alignVertical)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(alignHorizontal, alignVertical);
			Invoker.Method(this, "Align", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="distribute">NetOffice.VisioApi.Enums.VisDistributeTypes Distribute</param>
		/// <param name="glueToGuide">optional bool GlueToGuide = false</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Distribute(NetOffice.VisioApi.Enums.VisDistributeTypes distribute, object glueToGuide)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(distribute, glueToGuide);
			Invoker.Method(this, "Distribute", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="distribute">NetOffice.VisioApi.Enums.VisDistributeTypes Distribute</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Distribute(NetOffice.VisioApi.Enums.VisDistributeTypes distribute)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(distribute);
			Invoker.Method(this, "Distribute", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void UpdateAlignmentBox()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UpdateAlignmentBox", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="distance">Double Distance</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Offset(Double distance)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(distance);
			Invoker.Method(this, "Offset", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void ConnectShapes()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ConnectShapes", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection FlipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		/// <param name="pinY">optional Double PinY = 0</param>
		/// <param name="pinUnitsNameOrCode">optional object PinUnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType, object blastGuards, object pinX, object pinY, object pinUnitsNameOrCode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(flipDirection, flipType, blastGuards, pinX, pinY, pinUnitsNameOrCode);
			Invoker.Method(this, "Flip", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection FlipDirection</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(flipDirection);
			Invoker.Method(this, "Flip", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection FlipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(flipDirection, flipType);
			Invoker.Method(this, "Flip", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection FlipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType, object blastGuards)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(flipDirection, flipType, blastGuards);
			Invoker.Method(this, "Flip", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection FlipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType, object blastGuards, object pinX)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(flipDirection, flipType, blastGuards, pinX);
			Invoker.Method(this, "Flip", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="flipDirection">NetOffice.VisioApi.Enums.VisFlipDirection FlipDirection</param>
		/// <param name="flipType">optional NetOffice.VisioApi.Enums.VisFlipTypes FlipType = 0</param>
		/// <param name="blastGuards">optional bool BlastGuards = false</param>
		/// <param name="pinX">optional Double PinX = 0</param>
		/// <param name="pinY">optional Double PinY = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Flip(NetOffice.VisioApi.Enums.VisFlipDirection flipDirection, object flipType, object blastGuards, object pinX, object pinY)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(flipDirection, flipType, blastGuards, pinX, pinY);
			Invoker.Method(this, "Flip", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dataRecordsetID">Int32 DataRecordsetID</param>
		/// <param name="dataRowID">Int32 DataRowID</param>
		/// <param name="applyDataGraphicAfterLink">optional bool ApplyDataGraphicAfterLink = true</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void LinkToData(Int32 dataRecordsetID, Int32 dataRowID, object applyDataGraphicAfterLink)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID, dataRowID, applyDataGraphicAfterLink);
			Invoker.Method(this, "LinkToData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dataRecordsetID">Int32 DataRecordsetID</param>
		/// <param name="dataRowID">Int32 DataRowID</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void LinkToData(Int32 dataRecordsetID, Int32 dataRowID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID, dataRowID);
			Invoker.Method(this, "LinkToData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dataRecordsetID">Int32 DataRecordsetID</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void BreakLinkToData(Int32 dataRecordsetID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID);
			Invoker.Method(this, "BreakLinkToData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="shapeIDs">Int32[] ShapeIDs</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void GetIDs(out Int32[] shapeIDs)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			shapeIDs = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)shapeIDs);
			Invoker.Method(this, "GetIDs", paramsArray, modifiers);
			shapeIDs = (Int32[])paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dataRecordsetID">Int32 DataRecordsetID</param>
		/// <param name="columnNames">String[] ColumnNames</param>
		/// <param name="autoLinkFieldTypes">Int32[] AutoLinkFieldTypes</param>
		/// <param name="fieldNames">String[] FieldNames</param>
		/// <param name="autoLinkBehavior">Int32 AutoLinkBehavior</param>
		/// <param name="shapeIDs">Int32[] ShapeIDs</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public Int32 AutomaticLink(Int32 dataRecordsetID, String[] columnNames, Int32[] autoLinkFieldTypes, String[] fieldNames, Int32 autoLinkBehavior, out Int32[] shapeIDs)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,true);
			shapeIDs = null;
			object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID, (object)columnNames, (object)autoLinkFieldTypes, (object)fieldNames, autoLinkBehavior, (object)shapeIDs);
			object returnItem = Invoker.MethodReturn(this, "AutomaticLink", paramsArray);
			shapeIDs = (Int32[])paramsArray[5];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="alignOrSpace">NetOffice.VisioApi.Enums.VisLayoutIncrementalType AlignOrSpace</param>
		/// <param name="alignHorizontal">NetOffice.VisioApi.Enums.VisLayoutHorzAlignType AlignHorizontal</param>
		/// <param name="alignVertical">NetOffice.VisioApi.Enums.VisLayoutVertAlignType AlignVertical</param>
		/// <param name="spaceHorizontal">Double SpaceHorizontal</param>
		/// <param name="spaceVertical">Double SpaceVertical</param>
		/// <param name="unitCode">NetOffice.VisioApi.Enums.VisUnitCodes UnitCode</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void LayoutIncremental(NetOffice.VisioApi.Enums.VisLayoutIncrementalType alignOrSpace, NetOffice.VisioApi.Enums.VisLayoutHorzAlignType alignHorizontal, NetOffice.VisioApi.Enums.VisLayoutVertAlignType alignVertical, Double spaceHorizontal, Double spaceVertical, NetOffice.VisioApi.Enums.VisUnitCodes unitCode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(alignOrSpace, alignHorizontal, alignVertical, spaceHorizontal, spaceVertical, unitCode);
			Invoker.Method(this, "LayoutIncremental", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="direction">NetOffice.VisioApi.Enums.VisLayoutDirection Direction</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void LayoutChangeDirection(NetOffice.VisioApi.Enums.VisLayoutDirection direction)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(direction);
			Invoker.Method(this, "LayoutChangeDirection", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void AvoidPageBreaks()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AvoidPageBreaks", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="direction">NetOffice.VisioApi.Enums.VisResizeDirection Direction</param>
		/// <param name="distance">Double Distance</param>
		/// <param name="unitCode">NetOffice.VisioApi.Enums.VisUnitCodes UnitCode</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void Resize(NetOffice.VisioApi.Enums.VisResizeDirection direction, Double distance, NetOffice.VisioApi.Enums.VisUnitCodes unitCode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(direction, distance, unitCode);
			Invoker.Method(this, "Resize", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void AddToContainers()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AddToContainers", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void RemoveFromContainers()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RemoveFromContainers", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="page">NetOffice.VisioApi.IVPage Page</param>
		/// <param name="objectToDrop">object ObjectToDrop</param>
		/// <param name="newShape">optional NetOffice.VisioApi.IVShape NewShape = 0</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVSelection MoveToSubprocess(NetOffice.VisioApi.IVPage page, object objectToDrop, object newShape)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(page, objectToDrop, newShape);
			object returnItem = Invoker.MethodReturn(this, "MoveToSubprocess", paramsArray);
			NetOffice.VisioApi.IVSelection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVSelection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="page">NetOffice.VisioApi.IVPage Page</param>
		/// <param name="objectToDrop">object ObjectToDrop</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVSelection MoveToSubprocess(NetOffice.VisioApi.IVPage page, object objectToDrop)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(page, objectToDrop);
			object returnItem = Invoker.MethodReturn(this, "MoveToSubprocess", paramsArray);
			NetOffice.VisioApi.IVSelection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVSelection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="delFlags">Int32 DelFlags</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void DeleteEx(Int32 delFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(delFlags);
			Invoker.Method(this, "DeleteEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="nestedOptions">NetOffice.VisioApi.Enums.VisContainerNested NestedOptions</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32[] GetContainers(NetOffice.VisioApi.Enums.VisContainerNested nestedOptions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nestedOptions);
			object returnItem = (object)Invoker.MethodReturn(this, "GetContainers", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="nestedOptions">NetOffice.VisioApi.Enums.VisContainerNested NestedOptions</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32[] GetCallouts(NetOffice.VisioApi.Enums.VisContainerNested nestedOptions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nestedOptions);
			object returnItem = (object)Invoker.MethodReturn(this, "GetCallouts", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32[] MemberOfContainersUnion()
		{
			object[] paramsArray = null;
			object returnItem = (object)Invoker.MethodReturn(this, "MemberOfContainersUnion", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32[] MemberOfContainersIntersection()
		{
			object[] paramsArray = null;
			object returnItem = (object)Invoker.MethodReturn(this, "MemberOfContainersIntersection", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="formatType">NetOffice.VisioApi.Enums.VisContainerFormatType FormatType</param>
		/// <param name="formatValue">optional object FormatValue = 0</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32[] SetContainerFormat(NetOffice.VisioApi.Enums.VisContainerFormatType formatType, object formatValue)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formatType, formatValue);
			object returnItem = (object)Invoker.MethodReturn(this, "SetContainerFormat", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="formatType">NetOffice.VisioApi.Enums.VisContainerFormatType FormatType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32[] SetContainerFormat(NetOffice.VisioApi.Enums.VisContainerFormatType formatType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formatType);
			object returnItem = (object)Invoker.MethodReturn(this, "SetContainerFormat", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		/// <param name="masterOrMasterShortcutToDrop">object MasterOrMasterShortcutToDrop</param>
		/// <param name="replaceFlags">optional Int32 ReplaceFlags = 0</param>
		[SupportByVersionAttribute("Visio", 15, 16)]
		public NetOffice.VisioApi.IVShape[] ReplaceShape(object masterOrMasterShortcutToDrop, object replaceFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(masterOrMasterShortcutToDrop, replaceFlags);
			object returnItem = Invoker.MethodReturn(this, "ReplaceShape", paramsArray);
            ICOMObject[] newObject = Factory.CreateObjectArrayFromComProxy(this, (object[])returnItem);
			NetOffice.VisioApi.IVShape[] returnArray = new NetOffice.VisioApi.IVShape[newObject.Length];
			for (int i = 0; i < newObject.Length; i++)
				returnArray[i] = newObject[i] as NetOffice.VisioApi.IVShape;
			return returnArray;
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		/// <param name="masterOrMasterShortcutToDrop">object MasterOrMasterShortcutToDrop</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 15, 16)]
		public NetOffice.VisioApi.IVShape[] ReplaceShape(object masterOrMasterShortcutToDrop)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(masterOrMasterShortcutToDrop);
			object returnItem = Invoker.MethodReturn(this, "ReplaceShape", paramsArray);
            ICOMObject[] newObject = Factory.CreateObjectArrayFromComProxy(this, (object[])returnItem);
			NetOffice.VisioApi.IVShape[] returnArray = new NetOffice.VisioApi.IVShape[newObject.Length];
			for (int i = 0; i < newObject.Length; i++)
				returnArray[i] = newObject[i] as NetOffice.VisioApi.IVShape;
			return returnArray;
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		/// <param name="lineMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices lineMatrix</param>
		/// <param name="fillMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fillMatrix</param>
		/// <param name="effectsMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices effectsMatrix</param>
		/// <param name="fontMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fontMatrix</param>
		/// <param name="lineColor">NetOffice.VisioApi.Enums.VisQuickStyleColors lineColor</param>
		/// <param name="fillColor">NetOffice.VisioApi.Enums.VisQuickStyleColors fillColor</param>
		/// <param name="shadowColor">NetOffice.VisioApi.Enums.VisQuickStyleColors shadowColor</param>
		/// <param name="fontColor">NetOffice.VisioApi.Enums.VisQuickStyleColors fontColor</param>
		[SupportByVersionAttribute("Visio", 15, 16)]
		public void SetQuickStyle(NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices lineMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fillMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices effectsMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fontMatrix, NetOffice.VisioApi.Enums.VisQuickStyleColors lineColor, NetOffice.VisioApi.Enums.VisQuickStyleColors fillColor, NetOffice.VisioApi.Enums.VisQuickStyleColors shadowColor, NetOffice.VisioApi.Enums.VisQuickStyleColors fontColor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lineMatrix, fillMatrix, effectsMatrix, fontMatrix, lineColor, fillColor, shadowColor, fontColor);
			Invoker.Method(this, "SetQuickStyle", paramsArray);
		}

		#endregion

       #region IEnumerable<NetOffice.VisioApi.IVShape> Member
        
        /// <summary>
		/// SupportByVersionAttribute Visio, 11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
       public IEnumerator<NetOffice.VisioApi.IVShape> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.VisioApi.IVShape item in innerEnumerator)
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