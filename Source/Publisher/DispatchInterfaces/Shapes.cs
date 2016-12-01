using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.PublisherApi
{
	///<summary>
	/// DispatchInterface Shapes 
	/// SupportByVersion Publisher, 14,15,16
	///</summary>
	[SupportByVersionAttribute("Publisher", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Shapes : COMObject ,IEnumerable<NetOffice.PublisherApi.Shape>
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
                    _type = typeof(Shapes);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Shapes(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Shapes(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Shapes(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Shapes(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Shapes(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Shapes() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Shapes(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.PublisherApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Application.LateBindingApiWrapperType) as NetOffice.PublisherApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
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
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
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
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.pbCanvasArrangementType CanvasArrangementType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CanvasArrangementType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PublisherApi.Enums.pbCanvasArrangementType)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CanvasArrangementType", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public Int32 CanvasesCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CanvasesCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.PublisherApi.Shape this[object index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoCalloutType Type</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">object Width</param>
		/// <param name="height">object Height</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddCallout(NetOffice.OfficeApi.Enums.MsoCalloutType type, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddCallout", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoConnectorType Type</param>
		/// <param name="beginX">object BeginX</param>
		/// <param name="beginY">object BeginY</param>
		/// <param name="endX">object EndX</param>
		/// <param name="endY">object EndY</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddConnector(NetOffice.OfficeApi.Enums.MsoConnectorType type, object beginX, object beginY, object endX, object endY)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, beginX, beginY, endX, endY);
			object returnItem = Invoker.MethodReturn(this, "AddConnector", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="safeArrayOfPoints">object SafeArrayOfPoints</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddCurve(object safeArrayOfPoints)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(safeArrayOfPoints);
			object returnItem = Invoker.MethodReturn(this, "AddCurve", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="orientation">NetOffice.PublisherApi.Enums.PbTextOrientation Orientation</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">object Width</param>
		/// <param name="height">object Height</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddLabel(NetOffice.PublisherApi.Enums.PbTextOrientation orientation, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(orientation, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddLabel", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="beginX">object BeginX</param>
		/// <param name="beginY">object BeginY</param>
		/// <param name="endX">object EndX</param>
		/// <param name="endY">object EndY</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddLine(object beginX, object beginY, object endX, object endY)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(beginX, beginY, endX, endY);
			object returnItem = Invoker.MethodReturn(this, "AddLine", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		/// <param name="filename">optional string Filename = </param>
		/// <param name="link">optional NetOffice.OfficeApi.Enums.MsoTriState Link = -1</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object filename, object link)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height, className, filename, link);
			object returnItem = Invoker.MethodReturn(this, "AddOLEObject", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddOLEObject(object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top);
			object returnItem = Invoker.MethodReturn(this, "AddOLEObject", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">optional object Width = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width);
			object returnItem = Invoker.MethodReturn(this, "AddOLEObject", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddOLEObject", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width, object height, object className)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height, className);
			object returnItem = Invoker.MethodReturn(this, "AddOLEObject", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		/// <param name="filename">optional string Filename = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height, className, filename);
			object returnItem = Invoker.MethodReturn(this, "AddOLEObject", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState LinkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, linkToFile, saveWithDocument, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddPicture", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState LinkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, linkToFile, saveWithDocument, left, top);
			object returnItem = Invoker.MethodReturn(this, "AddPicture", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState LinkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">optional object Width = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, object left, object top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, linkToFile, saveWithDocument, left, top, width);
			object returnItem = Invoker.MethodReturn(this, "AddPicture", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="safeArrayOfPoints">object SafeArrayOfPoints</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddPolyline(object safeArrayOfPoints)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(safeArrayOfPoints);
			object returnItem = Invoker.MethodReturn(this, "AddPolyline", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoAutoShapeType Type</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">object Width</param>
		/// <param name="height">object Height</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddShape(NetOffice.OfficeApi.Enums.MsoAutoShapeType type, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddShape", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="numRows">Int32 NumRows</param>
		/// <param name="numColumns">Int32 NumColumns</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">object Width</param>
		/// <param name="height">object Height</param>
		/// <param name="fixedSize">optional bool FixedSize = false</param>
		/// <param name="direction">optional NetOffice.PublisherApi.Enums.PbTableDirectionType Direction = -1</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left, object top, object width, object height, object fixedSize, object direction)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numRows, numColumns, left, top, width, height, fixedSize, direction);
			object returnItem = Invoker.MethodReturn(this, "AddTable", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="numRows">Int32 NumRows</param>
		/// <param name="numColumns">Int32 NumColumns</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">object Width</param>
		/// <param name="height">object Height</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numRows, numColumns, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddTable", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="numRows">Int32 NumRows</param>
		/// <param name="numColumns">Int32 NumColumns</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">object Width</param>
		/// <param name="height">object Height</param>
		/// <param name="fixedSize">optional bool FixedSize = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left, object top, object width, object height, object fixedSize)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numRows, numColumns, left, top, width, height, fixedSize);
			object returnItem = Invoker.MethodReturn(this, "AddTable", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="orientation">NetOffice.PublisherApi.Enums.PbTextOrientation Orientation</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">object Width</param>
		/// <param name="height">object Height</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddTextbox(NetOffice.PublisherApi.Enums.PbTextOrientation orientation, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(orientation, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddTextbox", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="presetTextEffect">NetOffice.OfficeApi.Enums.MsoPresetTextEffect PresetTextEffect</param>
		/// <param name="text">string Text</param>
		/// <param name="fontName">string FontName</param>
		/// <param name="fontSize">object FontSize</param>
		/// <param name="fontBold">NetOffice.OfficeApi.Enums.MsoTriState FontBold</param>
		/// <param name="fontItalic">NetOffice.OfficeApi.Enums.MsoTriState FontItalic</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddTextEffect(NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect, string text, string fontName, object fontSize, NetOffice.OfficeApi.Enums.MsoTriState fontBold, NetOffice.OfficeApi.Enums.MsoTriState fontItalic, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(presetTextEffect, text, fontName, fontSize, fontBold, fontItalic, left, top);
			object returnItem = Invoker.MethodReturn(this, "AddTextEffect", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.PublisherApi.Enums.PbWebControlType Type</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">object Width</param>
		/// <param name="height">object Height</param>
		/// <param name="launchPropertiesWindow">optional bool LaunchPropertiesWindow = false</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddWebControl(NetOffice.PublisherApi.Enums.PbWebControlType type, object left, object top, object width, object height, object launchPropertiesWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, left, top, width, height, launchPropertiesWindow);
			object returnItem = Invoker.MethodReturn(this, "AddWebControl", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.PublisherApi.Enums.PbWebControlType Type</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">object Width</param>
		/// <param name="height">object Height</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddWebControl(NetOffice.PublisherApi.Enums.PbWebControlType type, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddWebControl", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType EditingType</param>
		/// <param name="x1">object X1</param>
		/// <param name="y1">object Y1</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.FreeformBuilder BuildFreeform(NetOffice.OfficeApi.Enums.MsoEditingType editingType, object x1, object y1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(editingType, x1, y1);
			object returnItem = Invoker.MethodReturn(this, "BuildFreeform", paramsArray);
			NetOffice.PublisherApi.FreeformBuilder newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.FreeformBuilder.LateBindingApiWrapperType) as NetOffice.PublisherApi.FreeformBuilder;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange Paste()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Paste", paramsArray);
			NetOffice.PublisherApi.ShapeRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.ShapeRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange Range(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "Range", paramsArray);
			NetOffice.PublisherApi.ShapeRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.ShapeRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange Range()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Range", paramsArray);
			NetOffice.PublisherApi.ShapeRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.ShapeRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void SelectAll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SelectAll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizardGroup Wizard</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		/// <param name="design">optional Int32 Design = -1</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddGroupWizard(NetOffice.PublisherApi.Enums.PbWizardGroup wizard, object left, object top, object width, object height, object design)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wizard, left, top, width, height, design);
			object returnItem = Invoker.MethodReturn(this, "AddGroupWizard", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizardGroup Wizard</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddGroupWizard(NetOffice.PublisherApi.Enums.PbWizardGroup wizard, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wizard, left, top);
			object returnItem = Invoker.MethodReturn(this, "AddGroupWizard", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizardGroup Wizard</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">optional object Width = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddGroupWizard(NetOffice.PublisherApi.Enums.PbWizardGroup wizard, object left, object top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wizard, left, top, width);
			object returnItem = Invoker.MethodReturn(this, "AddGroupWizard", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizardGroup Wizard</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddGroupWizard(NetOffice.PublisherApi.Enums.PbWizardGroup wizard, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wizard, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddGroupWizard", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wizardTag">NetOffice.PublisherApi.Enums.PbWizardTag WizardTag</param>
		/// <param name="instance">optional Int32 Instance = -1</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange FindShapeByWizardTag(NetOffice.PublisherApi.Enums.PbWizardTag wizardTag, object instance)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wizardTag, instance);
			object returnItem = Invoker.MethodReturn(this, "FindShapeByWizardTag", paramsArray);
			NetOffice.PublisherApi.ShapeRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.ShapeRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="wizardTag">NetOffice.PublisherApi.Enums.PbWizardTag WizardTag</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange FindShapeByWizardTag(NetOffice.PublisherApi.Enums.PbWizardTag wizardTag)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wizardTag);
			object returnItem = Invoker.MethodReturn(this, "FindShapeByWizardTag", paramsArray);
			NetOffice.PublisherApi.ShapeRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.ShapeRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">optional object Width</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddWebNavigationBar(string name, object left, object top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, left, top, width);
			object returnItem = Invoker.MethodReturn(this, "AddWebNavigationBar", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddWebNavigationBar(string name, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, left, top);
			object returnItem = Invoker.MethodReturn(this, "AddWebNavigationBar", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddCatalogMergeArea()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddCatalogMergeArea", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddEmptyPictureFrame(object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddEmptyPictureFrame", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddEmptyPictureFrame(object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top);
			object returnItem = Invoker.MethodReturn(this, "AddEmptyPictureFrame", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		/// <param name="width">optional object Width = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddEmptyPictureFrame(object left, object top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width);
			object returnItem = Invoker.MethodReturn(this, "AddEmptyPictureFrame", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="canvasId">Int32 CanvasId</param>
		/// <param name="catalogMergeFieldType">NetOffice.PublisherApi.Enums.pbCatalogMergeFieldType CatalogMergeFieldType</param>
		/// <param name="dbCol">Int32 DbCol</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void AddCatalogMergeFieldToCanvas(Int32 canvasId, NetOffice.PublisherApi.Enums.pbCatalogMergeFieldType catalogMergeFieldType, Int32 dbCol)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(canvasId, catalogMergeFieldType, dbCol);
			Invoker.Method(this, "AddCatalogMergeFieldToCanvas", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="presetWordArt">NetOffice.PublisherApi.Enums.pbPresetWordArt PresetWordArt</param>
		/// <param name="text">string Text</param>
		/// <param name="fontName">string FontName</param>
		/// <param name="fontSize">object FontSize</param>
		/// <param name="fontBold">NetOffice.OfficeApi.Enums.MsoTriState FontBold</param>
		/// <param name="fontItalic">NetOffice.OfficeApi.Enums.MsoTriState FontItalic</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddWordArt(NetOffice.PublisherApi.Enums.pbPresetWordArt presetWordArt, string text, string fontName, object fontSize, NetOffice.OfficeApi.Enums.MsoTriState fontBold, NetOffice.OfficeApi.Enums.MsoTriState fontItalic, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(presetWordArt, text, fontName, fontSize, fontBold, fontItalic, left, top);
			object returnItem = Invoker.MethodReturn(this, "AddWordArt", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bBlockIn">NetOffice.PublisherApi.BuildingBlock BBlockIn</param>
		/// <param name="left">object Left</param>
		/// <param name="top">object Top</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddBuildingBlock(NetOffice.PublisherApi.BuildingBlock bBlockIn, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bBlockIn, left, top);
			object returnItem = Invoker.MethodReturn(this, "AddBuildingBlock", paramsArray);
			NetOffice.PublisherApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PublisherApi.Shape.LateBindingApiWrapperType) as NetOffice.PublisherApi.Shape;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.PublisherApi.Shape> Member
        
        /// <summary>
		/// SupportByVersionAttribute Publisher, 14,15,16
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
       public IEnumerator<NetOffice.PublisherApi.Shape> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.PublisherApi.Shape item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Publisher, 14,15,16
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}