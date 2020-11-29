﻿using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface LPVISIOPAGE 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class LPVISIOPAGE : COMObject
	{
		#pragma warning disable

		#region Type Information

		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
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
                    _type = typeof(LPVISIOPAGE);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public LPVISIOPAGE(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIOPAGE(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOPAGE(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOPAGE(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOPAGE(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOPAGE(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOPAGE() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOPAGE(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Stat
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Background
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Background");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Background", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Index
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Index");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Index", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShapes Shapes
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShapes>(this, "Shapes");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVPage BackPageAsObj
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVPage>(this, "BackPageAsObj");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BackPageFromName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BackPageFromName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackPageFromName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVLayers Layers
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVLayers>(this, "Layers");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape PageSheet
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "PageSheet");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVEventList>(this, "EventList");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 PersistsEvents
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "PersistsEvents");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVConnects Connects
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVConnects>(this, "Connects");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public object BackPage
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BackPage");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BackPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 ID16
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ID16");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVOLEObjects OLEObjects
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVOLEObjects>(this, "OLEObjects");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 ID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="relation">Int16 relation</param>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVSelection get_SpatialSearch(Double x, Double y, Int16 relation, Double tolerance, Int16 flags)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVSelection>(this, "SpatialSearch", NetOffice.VisioApi.IVSelection.LateBindingApiWrapperType, new object[]{ x, y, relation, tolerance, flags });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_SpatialSearch
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="relation">Int16 relation</param>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_SpatialSearch")]
		public NetOffice.VisioApi.IVSelection SpatialSearch(Double x, Double y, Int16 relation, Double tolerance, Int16 flags)
		{
			return get_SpatialSearch(x, y, relation, tolerance, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string NameU
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NameU");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NameU", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), NativeResult]
		public stdole.Picture Picture
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Picture", paramsArray);
                return returnItem as stdole.Picture;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 PrintTileCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PrintTileCount");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisPageTypes Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisPageTypes>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 ReviewerID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ReviewerID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVPage OriginalPage
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVPage>(this, "OriginalPage");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public object ThemeColors
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ThemeColors");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ThemeColors", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public object ThemeEffects
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ThemeEffects");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ThemeEffects", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public bool LayoutRoutePassive
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LayoutRoutePassive");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LayoutRoutePassive", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public bool AutoSize
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		[BaseResult]
		public NetOffice.VisioApi.IVComments Comments
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVComments>(this, "Comments");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		[BaseResult]
		public NetOffice.VisioApi.IVComments ShapeComments
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVComments>(this, "ShapeComments");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void old_Paste()
		{
			 Factory.ExecuteMethod(this, "old_Paste");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">Int16 format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void old_PasteSpecial(Int16 format)
		{
			 Factory.ExecuteMethod(this, "old_PasteSpecial", format);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xBegin">Double xBegin</param>
		/// <param name="yBegin">Double yBegin</param>
		/// <param name="xEnd">Double xEnd</param>
		/// <param name="yEnd">Double yEnd</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DrawLine(Double xBegin, Double yBegin, Double xEnd, Double yEnd)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawLine", xBegin, yBegin, xEnd, yEnd);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="x1">Double x1</param>
		/// <param name="y1">Double y1</param>
		/// <param name="x2">Double x2</param>
		/// <param name="y2">Double y2</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DrawRectangle(Double x1, Double y1, Double x2, Double y2)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawRectangle", x1, y1, x2, y2);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="x1">Double x1</param>
		/// <param name="y1">Double y1</param>
		/// <param name="x2">Double x2</param>
		/// <param name="y2">Double y2</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DrawOval(Double x1, Double y1, Double x2, Double y2)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawOval", x1, y1, x2, y2);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape Drop(object objectToDrop, Double xPos, Double yPos)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "Drop", objectToDrop, xPos, yPos);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">Int16 type</param>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape AddGuide(Int16 type, Double xPos, Double yPos)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "AddGuide", type, xPos, yPos);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Print()
		{
			 Factory.ExecuteMethod(this, "Print");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape Import(string fileName)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "Import", fileName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Export(string fileName)
		{
			 Factory.ExecuteMethod(this, "Export", fileName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fRenumberPages">Int16 fRenumberPages</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Delete(Int16 fRenumberPages)
		{
			 Factory.ExecuteMethod(this, "Delete", fRenumberPages);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void CenterDrawing()
		{
			 Factory.ExecuteMethod(this, "CenterDrawing");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawSpline(Double[] xyArray, Double tolerance, Int16 flags)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)xyArray, tolerance, flags);
            object returnItem = Invoker.MethodReturn(this, "DrawSpline", paramsArray);
            NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this, returnItem, false) as NetOffice.VisioApi.IVShape;
            return newObject;
        }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="degree">Int16 degree</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawBezier(Double[] xyArray, Int16 degree, Int16 flags)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)xyArray, degree, flags);
            object returnItem = Invoker.MethodReturn(this, "DrawBezier", paramsArray);
            NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this, returnItem, false) as NetOffice.VisioApi.IVShape;
            return newObject;
        }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawPolyline(Double[] xyArray, Int16 flags)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)xyArray, flags);
            object returnItem = Invoker.MethodReturn(this, "DrawPolyline", paramsArray);
            NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this, returnItem, false) as NetOffice.VisioApi.IVShape;
            return newObject;
        }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape InsertFromFile(string fileName, Int16 flags)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "InsertFromFile", fileName, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="classOrProgID">string classOrProgID</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape InsertObject(string classOrProgID, Int16 flags)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "InsertObject", classOrProgID, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVWindow OpenDrawWindow()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "OpenDrawWindow");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectsToInstance">object[] objectsToInstance</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="iDArray">Int16[] iDArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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
		/// </summary>
		/// <param name="sID_SRCStream">Int16[] sID_SRCStream</param>
		/// <param name="formulaArray">object[] formulaArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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
		/// </summary>
		/// <param name="sID_SRCStream">Int16[] sID_SRCStream</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="unitsNamesOrCodes">object[] unitsNamesOrCodes</param>
		/// <param name="resultArray">object[] resultArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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
		/// </summary>
		/// <param name="sID_SRCStream">Int16[] sID_SRCStream</param>
		/// <param name="formulaArray">object[] formulaArray</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 SetFormulas(Int16[] sID_SRCStream, object[] formulaArray, Int16 flags)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)sID_SRCStream, (object)formulaArray, flags);
            object returnItem = Invoker.MethodReturn(this, "SetFormulas", paramsArray);
            return NetRuntimeSystem.Convert.ToInt16(returnItem);
        }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sID_SRCStream">Int16[] sID_SRCStream</param>
		/// <param name="unitsNamesOrCodes">object[] unitsNamesOrCodes</param>
		/// <param name="resultArray">object[] resultArray</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 SetResults(Int16[] sID_SRCStream, object[] unitsNamesOrCodes, object[] resultArray, Int16 flags)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)sID_SRCStream, (object)unitsNamesOrCodes, (object)resultArray, flags);
            object returnItem = Invoker.MethodReturn(this, "SetResults", paramsArray);
            return NetRuntimeSystem.Convert.ToInt16(returnItem);
        }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Layout()
		{
			 Factory.ExecuteMethod(this, "Layout");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flags">Int16 flags</param>
		/// <param name="lpr8Left">Double lpr8Left</param>
		/// <param name="lpr8Bottom">Double lpr8Bottom</param>
		/// <param name="lpr8Right">Double lpr8Right</param>
		/// <param name="lpr8Top">Double lpr8Top</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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
		/// </summary>
		/// <param name="objectsToInstance">object[] objectsToInstance</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="iDArray">Int16[] iDArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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
		/// </summary>
		/// <param name="sID_SRCStream">Int16[] sID_SRCStream</param>
		/// <param name="formulaArray">object[] formulaArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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
		/// </summary>
		/// <param name="degree">Int16 degree</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="knots">Double[] knots</param>
		/// <param name="weights">optional object weights</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawNURBS(Int16 degree, Int16 flags, Double[] xyArray, Double[] knots, object weights)
		{
            object[] paramsArray = Invoker.ValidateParamsArray(degree, flags, (object)xyArray, (object)knots, weights);
            object returnItem = Invoker.MethodReturn(this, "DrawNURBS", paramsArray);
            NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this, returnItem, false) as NetOffice.VisioApi.IVShape;
            return newObject;
        }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="degree">Int16 degree</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="knots">Double[] knots</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawNURBS(Int16 degree, Int16 flags, Double[] xyArray, Double[] knots)
		{
            object[] paramsArray = Invoker.ValidateParamsArray(degree, flags, (object)xyArray, (object)knots);
            object returnItem = Invoker.MethodReturn(this, "DrawNURBS", paramsArray);
            NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this, returnItem, false) as NetOffice.VisioApi.IVShape;
            return newObject;
        }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="nTile">Int32 nTile</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void PrintTile(Int32 nTile)
		{
			 Factory.ExecuteMethod(this, "PrintTile", nTile);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void ResizeToFitContents()
		{
			 Factory.ExecuteMethod(this, "ResizeToFitContents");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flags">optional object flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Paste(object flags)
		{
			 Factory.ExecuteMethod(this, "Paste", flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Paste()
		{
			 Factory.ExecuteMethod(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">Int32 format</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void PasteSpecial(Int32 format, object link, object displayAsIcon)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", format, link, displayAsIcon);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">Int32 format</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void PasteSpecial(Int32 format)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", format);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">Int32 format</param>
		/// <param name="link">optional object link</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void PasteSpecial(Int32 format, object link)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", format, link);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes selType</param>
		/// <param name="iterationMode">optional NetOffice.VisioApi.Enums.VisSelectMode IterationMode = 256</param>
		/// <param name="data">optional object data</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType, object iterationMode, object data)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "CreateSelection", selType, iterationMode, data);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes selType</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "CreateSelection", selType);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes selType</param>
		/// <param name="iterationMode">optional NetOffice.VisioApi.Enums.VisSelectMode IterationMode = 256</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType, object iterationMode)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "CreateSelection", selType, iterationMode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xBegin">Double xBegin</param>
		/// <param name="yBegin">Double yBegin</param>
		/// <param name="xEnd">Double xEnd</param>
		/// <param name="yEnd">Double yEnd</param>
		/// <param name="xControl">Double xControl</param>
		/// <param name="yControl">Double yControl</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DrawArcByThreePoints(Double xBegin, Double yBegin, Double xEnd, Double yEnd, Double xControl, Double yControl)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawArcByThreePoints", new object[]{ xBegin, yBegin, xEnd, yEnd, xControl, yControl });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xBegin">Double xBegin</param>
		/// <param name="yBegin">Double yBegin</param>
		/// <param name="xEnd">Double xEnd</param>
		/// <param name="yEnd">Double yEnd</param>
		/// <param name="sweepFlag">NetOffice.VisioApi.Enums.VisArcSweepFlags sweepFlag</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DrawQuarterArc(Double xBegin, Double yBegin, Double xEnd, Double yEnd, NetOffice.VisioApi.Enums.VisArcSweepFlags sweepFlag)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawQuarterArc", new object[]{ xBegin, yBegin, xEnd, yEnd, sweepFlag });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xCenter">Double xCenter</param>
		/// <param name="yCenter">Double yCenter</param>
		/// <param name="radius">Double radius</param>
		/// <param name="startAngle">optional Double StartAngle = 0</param>
		/// <param name="endAngle">optional Double EndAngle = 3.1415927410125732</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius, object startAngle, object endAngle)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawCircularArc", new object[]{ xCenter, yCenter, radius, startAngle, endAngle });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xCenter">Double xCenter</param>
		/// <param name="yCenter">Double yCenter</param>
		/// <param name="radius">Double radius</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawCircularArc", xCenter, yCenter, radius);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xCenter">Double xCenter</param>
		/// <param name="yCenter">Double yCenter</param>
		/// <param name="radius">Double radius</param>
		/// <param name="startAngle">optional Double StartAngle = 0</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius, object startAngle)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawCircularArc", xCenter, yCenter, radius, startAngle);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		[SupportByVersion("Visio", 12,14,15,16)]
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
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="dataRowID">Int32 dataRowID</param>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		[SupportByVersion("Visio", 12,14,15,16)]
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
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="dataRowIDs">Int32[] dataRowIDs</param>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		/// <param name="applyDataGraphicAfterLink">optional bool ApplyDataGraphicAfterLink = true</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public void LinkShapesToDataRows(Int32 dataRecordsetID, Int32[] dataRowIDs, Int32[] shapeIDs, object applyDataGraphicAfterLink)
		{
            object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID, (object)dataRowIDs, (object)shapeIDs, applyDataGraphicAfterLink);
            Invoker.Method(this, "LinkShapesToDataRows", paramsArray);
        }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="dataRowIDs">Int32[] dataRowIDs</param>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		[CustomMethod]
		[SupportByVersion("Visio", 12,14,15,16)]
		public void LinkShapesToDataRows(Int32 dataRecordsetID, Int32[] dataRowIDs, Int32[] shapeIDs)
		{
            object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID, (object)dataRowIDs, (object)shapeIDs);
            Invoker.Method(this, "LinkShapesToDataRows", paramsArray);
        }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		/// <param name="uniqueIDArgs">NetOffice.VisioApi.Enums.VisUniqueIDArgs uniqueIDArgs</param>
		/// <param name="gUIDs">String[] gUIDs</param>
		[SupportByVersion("Visio", 12,14,15,16)]
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
		/// </summary>
		/// <param name="gUIDs">String[] gUIDs</param>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		[SupportByVersion("Visio", 12,14,15,16)]
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
		/// </summary>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="dataRowID">Int32 dataRowID</param>
		/// <param name="applyDataGraphicAfterLink">bool applyDataGraphicAfterLink</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DropLinked(object objectToDrop, Double x, Double y, Int32 dataRecordsetID, Int32 dataRowID, bool applyDataGraphicAfterLink)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DropLinked", new object[]{ objectToDrop, x, y, dataRecordsetID, dataRowID, applyDataGraphicAfterLink });
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="objectsToInstance">object[] objectsToInstance</param>
		/// <param name="xYs">Double[] xYs</param>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="dataRowIDs">Int32[] dataRowIDs</param>
		/// <param name="applyDataGraphicAfterLink">bool applyDataGraphicAfterLink</param>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		[SupportByVersion("Visio", 12,14,15,16)]
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
		/// </summary>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="targetShape">NetOffice.VisioApi.IVShape targetShape</param>
		/// <param name="placementDir">NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir</param>
		/// <param name="connector">optional object Connector = null (Nothing in visual basic)</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DropConnected(object objectToDrop, NetOffice.VisioApi.IVShape targetShape, NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir, object connector)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DropConnected", objectToDrop, targetShape, placementDir, connector);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="targetShape">NetOffice.VisioApi.IVShape targetShape</param>
		/// <param name="placementDir">NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVShape DropConnected(object objectToDrop, NetOffice.VisioApi.IVShape targetShape, NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DropConnected", objectToDrop, targetShape, placementDir);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="fromShapeIDs">Int32[] fromShapeIDs</param>
		/// <param name="toShapeIDs">Int32[] toShapeIDs</param>
		/// <param name="placementDirs">Int32[] placementDirs</param>
		/// <param name="connector">optional object Connector = null (Nothing in visual basic)</param>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32 AutoConnectMany(Int32[] fromShapeIDs, Int32[] toShapeIDs, Int32[] placementDirs, object connector)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)fromShapeIDs, (object)toShapeIDs, (object)placementDirs, connector);
            object returnItem = Invoker.MethodReturn(this, "AutoConnectMany", paramsArray);
            return NetRuntimeSystem.Convert.ToInt32(returnItem);
        }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="fromShapeIDs">Int32[] fromShapeIDs</param>
		/// <param name="toShapeIDs">Int32[] toShapeIDs</param>
		/// <param name="placementDirs">Int32[] placementDirs</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		public Int32 AutoConnectMany(Int32[] fromShapeIDs, Int32[] toShapeIDs, Int32[] placementDirs)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)fromShapeIDs, (object)toShapeIDs, (object)placementDirs);
            object returnItem = Invoker.MethodReturn(this, "AutoConnectMany", paramsArray);
            return NetRuntimeSystem.Convert.ToInt32(returnItem);
        }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="targetShapes">object targetShapes</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DropContainer(object objectToDrop, object targetShapes)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DropContainer", objectToDrop, targetShapes);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="alignOrSpace">NetOffice.VisioApi.Enums.VisLayoutIncrementalType alignOrSpace</param>
		/// <param name="alignHorizontal">NetOffice.VisioApi.Enums.VisLayoutHorzAlignType alignHorizontal</param>
		/// <param name="alignVertical">NetOffice.VisioApi.Enums.VisLayoutVertAlignType alignVertical</param>
		/// <param name="spaceHorizontal">Double spaceHorizontal</param>
		/// <param name="spaceVertical">Double spaceVertical</param>
		/// <param name="unitsNameOrCode">NetOffice.VisioApi.Enums.VisUnitCodes unitsNameOrCode</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void LayoutIncremental(NetOffice.VisioApi.Enums.VisLayoutIncrementalType alignOrSpace, NetOffice.VisioApi.Enums.VisLayoutHorzAlignType alignHorizontal, NetOffice.VisioApi.Enums.VisLayoutVertAlignType alignVertical, Double spaceHorizontal, Double spaceVertical, NetOffice.VisioApi.Enums.VisUnitCodes unitsNameOrCode)
		{
			 Factory.ExecuteMethod(this, "LayoutIncremental", new object[]{ alignOrSpace, alignHorizontal, alignVertical, spaceHorizontal, spaceVertical, unitsNameOrCode });
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="direction">NetOffice.VisioApi.Enums.VisLayoutDirection direction</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void LayoutChangeDirection(NetOffice.VisioApi.Enums.VisLayoutDirection direction)
		{
			 Factory.ExecuteMethod(this, "LayoutChangeDirection", direction);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public void AvoidPageBreaks()
		{
			 Factory.ExecuteMethod(this, "AvoidPageBreaks");
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="connectorToSplit">NetOffice.VisioApi.IVShape connectorToSplit</param>
		/// <param name="shape">NetOffice.VisioApi.IVShape shape</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape SplitConnector(NetOffice.VisioApi.IVShape connectorToSplit, NetOffice.VisioApi.IVShape shape)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "SplitConnector", connectorToSplit, shape);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="targetShape">NetOffice.VisioApi.IVShape targetShape</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DropCallout(object objectToDrop, NetOffice.VisioApi.IVShape targetShape)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DropCallout", objectToDrop, targetShape);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		/// <param name="flags">Int32 flags</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void PasteToLocation(Double xPos, Double yPos, Int32 flags)
		{
			 Factory.ExecuteMethod(this, "PasteToLocation", xPos, yPos, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="nestedOptions">NetOffice.VisioApi.Enums.VisContainerNested nestedOptions</param>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32[] GetContainers(NetOffice.VisioApi.Enums.VisContainerNested nestedOptions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nestedOptions);
			object returnItem = (object)Invoker.MethodReturn(this, "GetContainers", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="nestedOptions">NetOffice.VisioApi.Enums.VisContainerNested nestedOptions</param>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32[] GetCallouts(NetOffice.VisioApi.Enums.VisContainerNested nestedOptions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nestedOptions);
			object returnItem = (object)Invoker.MethodReturn(this, "GetCallouts", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="outerList">object outerList</param>
		/// <param name="innerContainer">object innerContainer</param>
		/// <param name="populateFlags">NetOffice.VisioApi.Enums.VisLegendFlags populateFlags</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DropLegend(object outerList, object innerContainer, NetOffice.VisioApi.Enums.VisLegendFlags populateFlags)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DropLegend", outerList, innerContainer, populateFlags);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="targetList">NetOffice.VisioApi.IVShape targetList</param>
		/// <param name="lPosition">Int32 lPosition</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DropIntoList(object objectToDrop, NetOffice.VisioApi.IVShape targetList, Int32 lPosition)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DropIntoList", objectToDrop, targetList, lPosition);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public void AutoSizeDrawing()
		{
			 Factory.ExecuteMethod(this, "AutoSizeDrawing");
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		[BaseResult]
		public NetOffice.VisioApi.IVPage Duplicate()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVPage>(this, "Duplicate");
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="eThemeType">NetOffice.VisioApi.Enums.VisThemeTypes eThemeType</param>
		[SupportByVersion("Visio", 15, 16)]
		public object GetTheme(NetOffice.VisioApi.Enums.VisThemeTypes eThemeType)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetTheme", eThemeType);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="varThemeIndex">object varThemeIndex</param>
		/// <param name="varColorScheme">optional object varColorScheme</param>
		/// <param name="varEffectScheme">optional object varEffectScheme</param>
		/// <param name="varConnectorScheme">optional object varConnectorScheme</param>
		/// <param name="varFontScheme">optional object varFontScheme</param>
		[SupportByVersion("Visio", 15, 16)]
		public void SetTheme(object varThemeIndex, object varColorScheme, object varEffectScheme, object varConnectorScheme, object varFontScheme)
		{
			 Factory.ExecuteMethod(this, "SetTheme", new object[]{ varThemeIndex, varColorScheme, varEffectScheme, varConnectorScheme, varFontScheme });
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="varThemeIndex">object varThemeIndex</param>
		[CustomMethod]
		[SupportByVersion("Visio", 15, 16)]
		public void SetTheme(object varThemeIndex)
		{
			 Factory.ExecuteMethod(this, "SetTheme", varThemeIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="varThemeIndex">object varThemeIndex</param>
		/// <param name="varColorScheme">optional object varColorScheme</param>
		[CustomMethod]
		[SupportByVersion("Visio", 15, 16)]
		public void SetTheme(object varThemeIndex, object varColorScheme)
		{
			 Factory.ExecuteMethod(this, "SetTheme", varThemeIndex, varColorScheme);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="varThemeIndex">object varThemeIndex</param>
		/// <param name="varColorScheme">optional object varColorScheme</param>
		/// <param name="varEffectScheme">optional object varEffectScheme</param>
		[CustomMethod]
		[SupportByVersion("Visio", 15, 16)]
		public void SetTheme(object varThemeIndex, object varColorScheme, object varEffectScheme)
		{
			 Factory.ExecuteMethod(this, "SetTheme", varThemeIndex, varColorScheme, varEffectScheme);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="varThemeIndex">object varThemeIndex</param>
		/// <param name="varColorScheme">optional object varColorScheme</param>
		/// <param name="varEffectScheme">optional object varEffectScheme</param>
		/// <param name="varConnectorScheme">optional object varConnectorScheme</param>
		[CustomMethod]
		[SupportByVersion("Visio", 15, 16)]
		public void SetTheme(object varThemeIndex, object varColorScheme, object varEffectScheme, object varConnectorScheme)
		{
			 Factory.ExecuteMethod(this, "SetTheme", varThemeIndex, varColorScheme, varEffectScheme, varConnectorScheme);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="pVariantColor">Int16 pVariantColor</param>
		/// <param name="pVariantStyle">Int16 pVariantStyle</param>
		/// <param name="pEmbellishment">optional Int16 pEmbellishment = 0</param>
		[SupportByVersion("Visio", 15, 16)]
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
		/// </summary>
		/// <param name="pVariantColor">Int16 pVariantColor</param>
		/// <param name="pVariantStyle">Int16 pVariantStyle</param>
		[CustomMethod]
		[SupportByVersion("Visio", 15, 16)]
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
		/// </summary>
		/// <param name="variantColor">Int16 variantColor</param>
		/// <param name="variantStyle">Int16 variantStyle</param>
		/// <param name="embellishment">optional Int16 embellishment = -1</param>
		[SupportByVersion("Visio", 15, 16)]
		public void SetThemeVariant(Int16 variantColor, Int16 variantStyle, object embellishment)
		{
			 Factory.ExecuteMethod(this, "SetThemeVariant", variantColor, variantStyle, embellishment);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="variantColor">Int16 variantColor</param>
		/// <param name="variantStyle">Int16 variantStyle</param>
		[CustomMethod]
		[SupportByVersion("Visio", 15, 16)]
		public void SetThemeVariant(Int16 variantColor, Int16 variantStyle)
		{
			 Factory.ExecuteMethod(this, "SetThemeVariant", variantColor, variantStyle);
		}

		#endregion

		#pragma warning restore
	}
}
