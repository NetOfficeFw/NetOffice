using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.VisioApi
{
	///<summary>
	/// DispatchInterface IVPage 
	/// SupportByVersion Visio, 11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class IVPage : COMObject
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
                    _type = typeof(IVPage);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IVPage(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVPage(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVPage(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVPage(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVPage(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVPage() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVPage(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

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
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 Background
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Background", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Background", paramsArray);
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
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 Index
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Index", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Index", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Name", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShapes Shapes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shapes", paramsArray);
				NetOffice.VisioApi.IVShapes newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShapes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVPage BackPageAsObj
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BackPageAsObj", paramsArray);
				NetOffice.VisioApi.IVPage newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVPage;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BackPageFromName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BackPageFromName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "BackPageFromName", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVLayers Layers
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Layers", paramsArray);
				NetOffice.VisioApi.IVLayers newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVLayers;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape PageSheet
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PageSheet", paramsArray);
				NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
				return newObject;
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
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVConnects Connects
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Connects", paramsArray);
				NetOffice.VisioApi.IVConnects newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVConnects;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public object BackPage
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BackPage", paramsArray);
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
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "BackPage", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 ID16
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ID16", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVOLEObjects OLEObjects
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OLEObjects", paramsArray);
				NetOffice.VisioApi.IVOLEObjects newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVOLEObjects;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 ID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ID", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="relation">Int16 Relation</param>
		/// <param name="tolerance">Double Tolerance</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVSelection get_SpatialSearch(Double x, Double y, Int16 relation, Double tolerance, Int16 flags)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(x, y, relation, tolerance, flags);
			object returnItem = Invoker.PropertyGet(this, "SpatialSearch", paramsArray);
			NetOffice.VisioApi.IVSelection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVSelection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_SpatialSearch
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="relation">Int16 Relation</param>
		/// <param name="tolerance">Double Tolerance</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVSelection SpatialSearch(Double x, Double y, Int16 relation, Double tolerance, Int16 flags)
		{
			return get_SpatialSearch(x, y, relation, tolerance, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string NameU
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NameU", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "NameU", paramsArray);
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
		public Int32 PrintTileCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintTileCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisPageTypes Type
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Type", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.VisioApi.Enums.VisPageTypes)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 ReviewerID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReviewerID", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVPage OriginalPage
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OriginalPage", paramsArray);
				NetOffice.VisioApi.IVPage newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVPage;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public object ThemeColors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ThemeColors", paramsArray);
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
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ThemeColors", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public object ThemeEffects
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ThemeEffects", paramsArray);
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
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ThemeEffects", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public bool LayoutRoutePassive
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LayoutRoutePassive", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LayoutRoutePassive", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public bool AutoSize
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoSize", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AutoSize", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 15, 16)]
		public NetOffice.VisioApi.IVComments Comments
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Comments", paramsArray);
				NetOffice.VisioApi.IVComments newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVComments;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 15, 16)]
		public NetOffice.VisioApi.IVComments ShapeComments
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShapeComments", paramsArray);
				NetOffice.VisioApi.IVComments newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVComments;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void old_Paste()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "old_Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">Int16 Format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void old_PasteSpecial(Int16 format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format);
			Invoker.Method(this, "old_PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xBegin">Double xBegin</param>
		/// <param name="yBegin">Double yBegin</param>
		/// <param name="xEnd">Double xEnd</param>
		/// <param name="yEnd">Double yEnd</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawLine(Double xBegin, Double yBegin, Double xEnd, Double yEnd)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xBegin, yBegin, xEnd, yEnd);
			object returnItem = Invoker.MethodReturn(this, "DrawLine", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="x1">Double x1</param>
		/// <param name="y1">Double y1</param>
		/// <param name="x2">Double x2</param>
		/// <param name="y2">Double y2</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawRectangle(Double x1, Double y1, Double x2, Double y2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(x1, y1, x2, y2);
			object returnItem = Invoker.MethodReturn(this, "DrawRectangle", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="x1">Double x1</param>
		/// <param name="y1">Double y1</param>
		/// <param name="x2">Double x2</param>
		/// <param name="y2">Double y2</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawOval(Double x1, Double y1, Double x2, Double y2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(x1, y1, x2, y2);
			object returnItem = Invoker.MethodReturn(this, "DrawOval", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectToDrop">object ObjectToDrop</param>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape Drop(object objectToDrop, Double xPos, Double yPos)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectToDrop, xPos, yPos);
			object returnItem = Invoker.MethodReturn(this, "Drop", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">Int16 Type</param>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape AddGuide(Int16 type, Double xPos, Double yPos)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, xPos, yPos);
			object returnItem = Invoker.MethodReturn(this, "AddGuide", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Print()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Print", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape Import(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "Import", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

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
		/// <param name="fRenumberPages">Int16 fRenumberPages</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Delete(Int16 fRenumberPages)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fRenumberPages);
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void CenterDrawing()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CenterDrawing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="tolerance">Double Tolerance</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawSpline(Double[] xyArray, Double tolerance, Int16 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray((object)xyArray, tolerance, flags);
			object returnItem = Invoker.MethodReturn(this, "DrawSpline", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="degree">Int16 degree</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawBezier(Double[] xyArray, Int16 degree, Int16 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray((object)xyArray, degree, flags);
			object returnItem = Invoker.MethodReturn(this, "DrawBezier", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawPolyline(Double[] xyArray, Int16 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray((object)xyArray, flags);
			object returnItem = Invoker.MethodReturn(this, "DrawPolyline", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape InsertFromFile(string fileName, Int16 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, flags);
			object returnItem = Invoker.MethodReturn(this, "InsertFromFile", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="classOrProgID">string ClassOrProgID</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape InsertObject(string classOrProgID, Int16 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(classOrProgID, flags);
			object returnItem = Invoker.MethodReturn(this, "InsertObject", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow OpenDrawWindow()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "OpenDrawWindow", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectsToInstance">object[] ObjectsToInstance</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="iDArray">Int16[] IDArray</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 DropMany(object[] objectsToInstance, Double[] xyArray, out Int16[] iDArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			iDArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)objectsToInstance, (object)xyArray, (object)iDArray);
			object returnItem = Invoker.MethodReturn(this, "DropMany", paramsArray);
			iDArray = (Int16[])paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}       

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// 
        /// </summary>
        /// <param name="sID_SRCStream">Int16[] SID_SRCStream</param>
        /// <param name="formulaArray">object[] formulaArray</param>
        [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void GetFormulas(Int16[] sID_SRCStream, out object[] formulaArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			formulaArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)sID_SRCStream, (object)formulaArray);
			Invoker.Method(this, "GetFormulas", paramsArray, modifiers);
			formulaArray = (object[])paramsArray[1];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sID_SRCStream">Int16[] SID_SRCStream</param>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="unitsNamesOrCodes">object[] UnitsNamesOrCodes</param>
		/// <param name="resultArray">object[] resultArray</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void GetResults(Int16[] sID_SRCStream, Int16 flags, object[] unitsNamesOrCodes, out object[] resultArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true);
			resultArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)sID_SRCStream, flags, (object)unitsNamesOrCodes, (object)resultArray);
			Invoker.Method(this, "GetResults", paramsArray, modifiers);
			resultArray = (object[])paramsArray[3];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sID_SRCStream">Int16[] SID_SRCStream</param>
		/// <param name="formulaArray">object[] formulaArray</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 SetFormulas(Int16[] sID_SRCStream, object[] formulaArray, Int16 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray((object)sID_SRCStream, (object)formulaArray, flags);
			object returnItem = Invoker.MethodReturn(this, "SetFormulas", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sID_SRCStream">Int16[] SID_SRCStream</param>
		/// <param name="unitsNamesOrCodes">object[] UnitsNamesOrCodes</param>
		/// <param name="resultArray">object[] resultArray</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 SetResults(Int16[] sID_SRCStream, object[] unitsNamesOrCodes, object[] resultArray, Int16 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray((object)sID_SRCStream, (object)unitsNamesOrCodes, (object)resultArray, flags);
			object returnItem = Invoker.MethodReturn(this, "SetResults", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
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
		/// <param name="objectsToInstance">object[] ObjectsToInstance</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="iDArray">Int16[] IDArray</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 DropManyU(object[] objectsToInstance, Double[] xyArray, out Int16[] iDArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			iDArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)objectsToInstance, (object)xyArray, (object)iDArray);
			object returnItem = Invoker.MethodReturn(this, "DropManyU", paramsArray);
			iDArray = (Int16[])paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sID_SRCStream">Int16[] SID_SRCStream</param>
		/// <param name="formulaArray">object[] formulaArray</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void GetFormulasU(Int16[] sID_SRCStream, out object[] formulaArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			formulaArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)sID_SRCStream, (object)formulaArray);
			Invoker.Method(this, "GetFormulasU", paramsArray, modifiers);
			formulaArray = (object[])paramsArray[1];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="degree">Int16 degree</param>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="knots">Double[] knots</param>
		/// <param name="weights">optional object weights</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawNURBS(Int16 degree, Int16 flags, Double[] xyArray, Double[] knots, object weights)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(degree, flags, (object)xyArray, (object)knots, weights);
			object returnItem = Invoker.MethodReturn(this, "DrawNURBS", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="degree">Int16 degree</param>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="knots">Double[] knots</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawNURBS(Int16 degree, Int16 flags, Double[] xyArray, Double[] knots)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(degree, flags, (object)xyArray, (object)knots);
			object returnItem = Invoker.MethodReturn(this, "DrawNURBS", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="nTile">Int32 nTile</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void PrintTile(Int32 nTile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nTile);
			Invoker.Method(this, "PrintTile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void ResizeToFitContents()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ResizeToFitContents", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="flags">optional object Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Paste(object flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(flags);
			Invoker.Method(this, "Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Paste()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">Int32 Format</param>
		/// <param name="link">optional object Link</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void PasteSpecial(Int32 format, object link, object displayAsIcon)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, link, displayAsIcon);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">Int32 Format</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void PasteSpecial(Int32 format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">Int32 Format</param>
		/// <param name="link">optional object Link</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void PasteSpecial(Int32 format, object link)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, link);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes SelType</param>
		/// <param name="iterationMode">optional NetOffice.VisioApi.Enums.VisSelectMode IterationMode = 256</param>
		/// <param name="data">optional object Data</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType, object iterationMode, object data)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(selType, iterationMode, data);
			object returnItem = Invoker.MethodReturn(this, "CreateSelection", paramsArray);
			NetOffice.VisioApi.IVSelection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVSelection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes SelType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(selType);
			object returnItem = Invoker.MethodReturn(this, "CreateSelection", paramsArray);
			NetOffice.VisioApi.IVSelection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVSelection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes SelType</param>
		/// <param name="iterationMode">optional NetOffice.VisioApi.Enums.VisSelectMode IterationMode = 256</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType, object iterationMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(selType, iterationMode);
			object returnItem = Invoker.MethodReturn(this, "CreateSelection", paramsArray);
			NetOffice.VisioApi.IVSelection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVSelection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xBegin">Double xBegin</param>
		/// <param name="yBegin">Double yBegin</param>
		/// <param name="xEnd">Double xEnd</param>
		/// <param name="yEnd">Double yEnd</param>
		/// <param name="xControl">Double xControl</param>
		/// <param name="yControl">Double yControl</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawArcByThreePoints(Double xBegin, Double yBegin, Double xEnd, Double yEnd, Double xControl, Double yControl)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xBegin, yBegin, xEnd, yEnd, xControl, yControl);
			object returnItem = Invoker.MethodReturn(this, "DrawArcByThreePoints", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xBegin">Double xBegin</param>
		/// <param name="yBegin">Double yBegin</param>
		/// <param name="xEnd">Double xEnd</param>
		/// <param name="yEnd">Double yEnd</param>
		/// <param name="sweepFlag">NetOffice.VisioApi.Enums.VisArcSweepFlags SweepFlag</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawQuarterArc(Double xBegin, Double yBegin, Double xEnd, Double yEnd, NetOffice.VisioApi.Enums.VisArcSweepFlags sweepFlag)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xBegin, yBegin, xEnd, yEnd, sweepFlag);
			object returnItem = Invoker.MethodReturn(this, "DrawQuarterArc", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xCenter">Double xCenter</param>
		/// <param name="yCenter">Double yCenter</param>
		/// <param name="radius">Double Radius</param>
		/// <param name="startAngle">optional Double StartAngle = 0</param>
		/// <param name="endAngle">optional Double EndAngle = 3.1415927410125732</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius, object startAngle, object endAngle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xCenter, yCenter, radius, startAngle, endAngle);
			object returnItem = Invoker.MethodReturn(this, "DrawCircularArc", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xCenter">Double xCenter</param>
		/// <param name="yCenter">Double yCenter</param>
		/// <param name="radius">Double Radius</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xCenter, yCenter, radius);
			object returnItem = Invoker.MethodReturn(this, "DrawCircularArc", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xCenter">Double xCenter</param>
		/// <param name="yCenter">Double yCenter</param>
		/// <param name="radius">Double Radius</param>
		/// <param name="startAngle">optional Double StartAngle = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius, object startAngle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xCenter, yCenter, radius, startAngle);
			object returnItem = Invoker.MethodReturn(this, "DrawCircularArc", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dataRecordsetID">Int32 DataRecordsetID</param>
		/// <param name="shapeIDs">Int32[] ShapeIDs</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void GetShapesLinkedToData(Int32 dataRecordsetID, out Int32[] shapeIDs)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			shapeIDs = null;
			object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID, (object)shapeIDs);
			Invoker.Method(this, "GetShapesLinkedToData", paramsArray, modifiers);
			shapeIDs = (Int32[])paramsArray[1];
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dataRecordsetID">Int32 DataRecordsetID</param>
		/// <param name="dataRowID">Int32 DataRowID</param>
		/// <param name="shapeIDs">Int32[] ShapeIDs</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void GetShapesLinkedToDataRow(Int32 dataRecordsetID, Int32 dataRowID, out Int32[] shapeIDs)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			shapeIDs = null;
			object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID, dataRowID, (object)shapeIDs);
			Invoker.Method(this, "GetShapesLinkedToDataRow", paramsArray, modifiers);
			shapeIDs = (Int32[])paramsArray[2];
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dataRecordsetID">Int32 DataRecordsetID</param>
		/// <param name="dataRowIDs">Int32[] DataRowIDs</param>
		/// <param name="shapeIDs">Int32[] ShapeIDs</param>
		/// <param name="applyDataGraphicAfterLink">optional bool ApplyDataGraphicAfterLink = true</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void LinkShapesToDataRows(Int32 dataRecordsetID, Int32[] dataRowIDs, Int32[] shapeIDs, object applyDataGraphicAfterLink)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID, (object)dataRowIDs, (object)shapeIDs, applyDataGraphicAfterLink);
			Invoker.Method(this, "LinkShapesToDataRows", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dataRecordsetID">Int32 DataRecordsetID</param>
		/// <param name="dataRowIDs">Int32[] DataRowIDs</param>
		/// <param name="shapeIDs">Int32[] ShapeIDs</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void LinkShapesToDataRows(Int32 dataRecordsetID, Int32[] dataRowIDs, Int32[] shapeIDs)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID, (object)dataRowIDs, (object)shapeIDs);
			Invoker.Method(this, "LinkShapesToDataRows", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="shapeIDs">Int32[] ShapeIDs</param>
		/// <param name="uniqueIDArgs">NetOffice.VisioApi.Enums.VisUniqueIDArgs UniqueIDArgs</param>
		/// <param name="gUIDs">String[] GUIDs</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void ShapeIDsToUniqueIDs(Int32[] shapeIDs, NetOffice.VisioApi.Enums.VisUniqueIDArgs uniqueIDArgs, out String[] gUIDs)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			gUIDs = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)shapeIDs, uniqueIDArgs, (object)gUIDs);
			Invoker.Method(this, "ShapeIDsToUniqueIDs", paramsArray, modifiers);
			gUIDs = (String[])paramsArray[2];
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="gUIDs">String[] GUIDs</param>
		/// <param name="shapeIDs">Int32[] ShapeIDs</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void UniqueIDsToShapeIDs(String[] gUIDs, out Int32[] shapeIDs)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			shapeIDs = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)gUIDs, (object)shapeIDs);
			Invoker.Method(this, "UniqueIDsToShapeIDs", paramsArray, modifiers);
			shapeIDs = (Int32[])paramsArray[1];
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectToDrop">object ObjectToDrop</param>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="dataRecordsetID">Int32 DataRecordsetID</param>
		/// <param name="dataRowID">Int32 DataRowID</param>
		/// <param name="applyDataGraphicAfterLink">bool ApplyDataGraphicAfterLink</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.IVShape DropLinked(object objectToDrop, Double x, Double y, Int32 dataRecordsetID, Int32 dataRowID, bool applyDataGraphicAfterLink)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectToDrop, x, y, dataRecordsetID, dataRowID, applyDataGraphicAfterLink);
			object returnItem = Invoker.MethodReturn(this, "DropLinked", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectsToInstance">object[] ObjectsToInstance</param>
		/// <param name="xYs">Double[] XYs</param>
		/// <param name="dataRecordsetID">Int32 DataRecordsetID</param>
		/// <param name="dataRowIDs">Int32[] DataRowIDs</param>
		/// <param name="applyDataGraphicAfterLink">bool ApplyDataGraphicAfterLink</param>
		/// <param name="shapeIDs">Int32[] ShapeIDs</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public Int32 DropManyLinkedU(object[] objectsToInstance, Double[] xYs, Int32 dataRecordsetID, Int32[] dataRowIDs, bool applyDataGraphicAfterLink, out Int32[] shapeIDs)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,true);
			shapeIDs = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)objectsToInstance, (object)xYs, dataRecordsetID, (object)dataRowIDs, applyDataGraphicAfterLink, (object)shapeIDs);
			object returnItem = Invoker.MethodReturn(this, "DropManyLinkedU", paramsArray);
			shapeIDs = (Int32[])paramsArray[5];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectToDrop">object ObjectToDrop</param>
		/// <param name="targetShape">NetOffice.VisioApi.IVShape TargetShape</param>
		/// <param name="placementDir">NetOffice.VisioApi.Enums.VisAutoConnectDir PlacementDir</param>
		/// <param name="connector">optional object Connector = null (Nothing in visual basic)</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVShape DropConnected(object objectToDrop, NetOffice.VisioApi.IVShape targetShape, NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir, object connector)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectToDrop, targetShape, placementDir, connector);
			object returnItem = Invoker.MethodReturn(this, "DropConnected", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectToDrop">object ObjectToDrop</param>
		/// <param name="targetShape">NetOffice.VisioApi.IVShape TargetShape</param>
		/// <param name="placementDir">NetOffice.VisioApi.Enums.VisAutoConnectDir PlacementDir</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVShape DropConnected(object objectToDrop, NetOffice.VisioApi.IVShape targetShape, NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectToDrop, targetShape, placementDir);
			object returnItem = Invoker.MethodReturn(this, "DropConnected", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fromShapeIDs">Int32[] FromShapeIDs</param>
		/// <param name="toShapeIDs">Int32[] ToShapeIDs</param>
		/// <param name="placementDirs">Int32[] PlacementDirs</param>
		/// <param name="connector">optional object Connector = null (Nothing in visual basic)</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32 AutoConnectMany(Int32[] fromShapeIDs, Int32[] toShapeIDs, Int32[] placementDirs, object connector)
		{
			object[] paramsArray = Invoker.ValidateParamsArray((object)fromShapeIDs, (object)toShapeIDs, (object)placementDirs, connector);
			object returnItem = Invoker.MethodReturn(this, "AutoConnectMany", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fromShapeIDs">Int32[] FromShapeIDs</param>
		/// <param name="toShapeIDs">Int32[] ToShapeIDs</param>
		/// <param name="placementDirs">Int32[] PlacementDirs</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32 AutoConnectMany(Int32[] fromShapeIDs, Int32[] toShapeIDs, Int32[] placementDirs)
		{
			object[] paramsArray = Invoker.ValidateParamsArray((object)fromShapeIDs, (object)toShapeIDs, (object)placementDirs);
			object returnItem = Invoker.MethodReturn(this, "AutoConnectMany", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectToDrop">object ObjectToDrop</param>
		/// <param name="targetShapes">object TargetShapes</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVShape DropContainer(object objectToDrop, object targetShapes)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectToDrop, targetShapes);
			object returnItem = Invoker.MethodReturn(this, "DropContainer", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
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
		/// <param name="unitsNameOrCode">NetOffice.VisioApi.Enums.VisUnitCodes UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void LayoutIncremental(NetOffice.VisioApi.Enums.VisLayoutIncrementalType alignOrSpace, NetOffice.VisioApi.Enums.VisLayoutHorzAlignType alignHorizontal, NetOffice.VisioApi.Enums.VisLayoutVertAlignType alignVertical, Double spaceHorizontal, Double spaceVertical, NetOffice.VisioApi.Enums.VisUnitCodes unitsNameOrCode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(alignOrSpace, alignHorizontal, alignVertical, spaceHorizontal, spaceVertical, unitsNameOrCode);
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
		/// <param name="connectorToSplit">NetOffice.VisioApi.IVShape ConnectorToSplit</param>
		/// <param name="shape">NetOffice.VisioApi.IVShape Shape</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVShape SplitConnector(NetOffice.VisioApi.IVShape connectorToSplit, NetOffice.VisioApi.IVShape shape)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(connectorToSplit, shape);
			object returnItem = Invoker.MethodReturn(this, "SplitConnector", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectToDrop">object ObjectToDrop</param>
		/// <param name="targetShape">NetOffice.VisioApi.IVShape TargetShape</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVShape DropCallout(object objectToDrop, NetOffice.VisioApi.IVShape targetShape)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectToDrop, targetShape);
			object returnItem = Invoker.MethodReturn(this, "DropCallout", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		/// <param name="flags">Int32 Flags</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void PasteToLocation(Double xPos, Double yPos, Int32 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPos, yPos, flags);
			Invoker.Method(this, "PasteToLocation", paramsArray);
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
		/// <param name="outerList">object OuterList</param>
		/// <param name="innerContainer">object InnerContainer</param>
		/// <param name="populateFlags">NetOffice.VisioApi.Enums.VisLegendFlags populateFlags</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVShape DropLegend(object outerList, object innerContainer, NetOffice.VisioApi.Enums.VisLegendFlags populateFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(outerList, innerContainer, populateFlags);
			object returnItem = Invoker.MethodReturn(this, "DropLegend", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectToDrop">object ObjectToDrop</param>
		/// <param name="targetList">NetOffice.VisioApi.IVShape TargetList</param>
		/// <param name="lPosition">Int32 lPosition</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVShape DropIntoList(object objectToDrop, NetOffice.VisioApi.IVShape targetList, Int32 lPosition)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectToDrop, targetList, lPosition);
			object returnItem = Invoker.MethodReturn(this, "DropIntoList", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void AutoSizeDrawing()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AutoSizeDrawing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 15, 16)]
		public NetOffice.VisioApi.IVPage Duplicate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Duplicate", paramsArray);
			NetOffice.VisioApi.IVPage newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVPage;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		/// <param name="eThemeType">NetOffice.VisioApi.Enums.VisThemeTypes eThemeType</param>
		[SupportByVersionAttribute("Visio", 15, 16)]
		public object GetTheme(NetOffice.VisioApi.Enums.VisThemeTypes eThemeType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(eThemeType);
			object returnItem = Invoker.MethodReturn(this, "GetTheme", paramsArray);
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
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		/// <param name="varThemeIndex">object varThemeIndex</param>
		/// <param name="varColorScheme">optional object varColorScheme</param>
		/// <param name="varEffectScheme">optional object varEffectScheme</param>
		/// <param name="varConnectorScheme">optional object varConnectorScheme</param>
		/// <param name="varFontScheme">optional object varFontScheme</param>
		[SupportByVersionAttribute("Visio", 15, 16)]
		public void SetTheme(object varThemeIndex, object varColorScheme, object varEffectScheme, object varConnectorScheme, object varFontScheme)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varThemeIndex, varColorScheme, varEffectScheme, varConnectorScheme, varFontScheme);
			Invoker.Method(this, "SetTheme", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		/// <param name="varThemeIndex">object varThemeIndex</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 15, 16)]
		public void SetTheme(object varThemeIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varThemeIndex);
			Invoker.Method(this, "SetTheme", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		/// <param name="varThemeIndex">object varThemeIndex</param>
		/// <param name="varColorScheme">optional object varColorScheme</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 15, 16)]
		public void SetTheme(object varThemeIndex, object varColorScheme)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varThemeIndex, varColorScheme);
			Invoker.Method(this, "SetTheme", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		/// <param name="varThemeIndex">object varThemeIndex</param>
		/// <param name="varColorScheme">optional object varColorScheme</param>
		/// <param name="varEffectScheme">optional object varEffectScheme</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 15, 16)]
		public void SetTheme(object varThemeIndex, object varColorScheme, object varEffectScheme)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varThemeIndex, varColorScheme, varEffectScheme);
			Invoker.Method(this, "SetTheme", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		/// <param name="varThemeIndex">object varThemeIndex</param>
		/// <param name="varColorScheme">optional object varColorScheme</param>
		/// <param name="varEffectScheme">optional object varEffectScheme</param>
		/// <param name="varConnectorScheme">optional object varConnectorScheme</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 15, 16)]
		public void SetTheme(object varThemeIndex, object varColorScheme, object varEffectScheme, object varConnectorScheme)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varThemeIndex, varColorScheme, varEffectScheme, varConnectorScheme);
			Invoker.Method(this, "SetTheme", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		/// <param name="pVariantColor">Int16 pVariantColor</param>
		/// <param name="pVariantStyle">Int16 pVariantStyle</param>
		/// <param name="pEmbellishment">optional Int16 pEmbellishment = 0</param>
		[SupportByVersionAttribute("Visio", 15, 16)]
		public void GetThemeVariant(out Int16 pVariantColor, out Int16 pVariantStyle, object pEmbellishment)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,false);
			pVariantColor = 0;
			pVariantStyle = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pVariantColor, pVariantStyle, pEmbellishment);
			Invoker.Method(this, "GetThemeVariant", paramsArray, modifiers);
			pVariantColor = (Int16)paramsArray[0];
			pVariantStyle = (Int16)paramsArray[1];
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		/// <param name="pVariantColor">Int16 pVariantColor</param>
		/// <param name="pVariantStyle">Int16 pVariantStyle</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 15, 16)]
		public void GetThemeVariant(out Int16 pVariantColor, out Int16 pVariantStyle)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true);
			pVariantColor = 0;
			pVariantStyle = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pVariantColor, pVariantStyle);
			Invoker.Method(this, "GetThemeVariant", paramsArray, modifiers);
			pVariantColor = (Int16)paramsArray[0];
			pVariantStyle = (Int16)paramsArray[1];
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		/// <param name="variantColor">Int16 variantColor</param>
		/// <param name="variantStyle">Int16 variantStyle</param>
		/// <param name="embellishment">optional Int16 embellishment = -1</param>
		[SupportByVersionAttribute("Visio", 15, 16)]
		public void SetThemeVariant(Int16 variantColor, Int16 variantStyle, object embellishment)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(variantColor, variantStyle, embellishment);
			Invoker.Method(this, "SetThemeVariant", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		/// <param name="variantColor">Int16 variantColor</param>
		/// <param name="variantStyle">Int16 variantStyle</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 15, 16)]
		public void SetThemeVariant(Int16 variantColor, Int16 variantStyle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(variantColor, variantStyle);
			Invoker.Method(this, "SetThemeVariant", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}