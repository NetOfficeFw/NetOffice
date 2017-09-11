using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVMaster 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IVMaster : COMObject
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
                    _type = typeof(IVMaster);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IVMaster(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IVMaster(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVMaster(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVMaster(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVMaster(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVMaster(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVMaster() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVMaster(string progId) : base(progId)
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string Prompt
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Prompt");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Prompt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 AlignName
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "AlignName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AlignName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 IconSize
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "IconSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IconSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 IconUpdate
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "IconUpdate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IconUpdate", value);
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
		public Int16 Index
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 OneD
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "OneD");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string UniqueID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UniqueID");
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 PatternFlags
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "PatternFlags");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PatternFlags", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 MatchByName
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "MatchByName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchByName", value);
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Hidden
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Hidden");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Hidden", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string BaseID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BaseID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string NewBaseID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NewBaseID");
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 IndexInStencil
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "IndexInStencil");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IndexInStencil", value);
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), NativeResult]
		public stdole.Picture Icon
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Icon", paramsArray);
                return returnItem as stdole.Picture;
            }
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Icon", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVMaster EditCopy
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVMaster>(this, "EditCopy");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVMaster Original
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVMaster>(this, "Original");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool IsChanged
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsChanged");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisMasterTypes Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisMasterTypes>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public bool DataGraphicHidden
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DataGraphicHidden");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataGraphicHidden", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public bool DataGraphicHidesText
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DataGraphicHidesText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataGraphicHidesText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public bool DataGraphicShowBorder
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DataGraphicShowBorder");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataGraphicShowBorder", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisGraphicPositionHorizontal DataGraphicHorizontalPosition
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisGraphicPositionHorizontal>(this, "DataGraphicHorizontalPosition");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DataGraphicHorizontalPosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisGraphicPositionVertical DataGraphicVerticalPosition
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisGraphicPositionVertical>(this, "DataGraphicVerticalPosition");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DataGraphicVerticalPosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVGraphicItems GraphicItems
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVGraphicItems>(this, "GraphicItems");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
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
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void CenterDrawing()
		{
			 Factory.ExecuteMethod(this, "CenterDrawing");
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
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVWindow OpenIconWindow()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "OpenIconWindow");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVMaster Open()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVMaster>(this, "Open");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Close()
		{
			 Factory.ExecuteMethod(this, "Close");
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
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void ImportIcon(string fileName)
		{
			 Factory.ExecuteMethod(this, "ImportIcon", fileName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="flags">Int16 flags</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void ExportIconTransparentAsBlack(string fileName, Int16 flags)
		{
			 Factory.ExecuteMethod(this, "ExportIconTransparentAsBlack", fileName, flags);
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
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVMasterShortcut CreateShortcut()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVMasterShortcut>(this, "CreateShortcut");
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
		/// <param name="fileName">string fileName</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="transparentRGB">optional object transparentRGB</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void ExportIcon(string fileName, Int16 flags, object transparentRGB)
		{
			 Factory.ExecuteMethod(this, "ExportIcon", fileName, flags, transparentRGB);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="flags">Int16 flags</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void ExportIcon(string fileName, Int16 flags)
		{
			 Factory.ExecuteMethod(this, "ExportIcon", fileName, flags);
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
		[SupportByVersion("Visio", 12,14,15,16)]
		public void DataGraphicDelete()
		{
			 Factory.ExecuteMethod(this, "DataGraphicDelete");
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

		#endregion

		#pragma warning restore
	}
}
