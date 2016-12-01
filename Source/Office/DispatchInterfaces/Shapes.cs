using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// DispatchInterface Shapes 
	/// SupportByVersion Office, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Shapes : _IMsoDispObj ,IEnumerable<NetOffice.OfficeApi.Shape>
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
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape Background
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Background", paramsArray);
				NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape Default
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Default", paramsArray);
				NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.OfficeApi.Shape this[object index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoCalloutType Type</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddCallout(NetOffice.OfficeApi.Enums.MsoCalloutType type, Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddCallout", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoConnectorType Type</param>
		/// <param name="beginX">Single BeginX</param>
		/// <param name="beginY">Single BeginY</param>
		/// <param name="endX">Single EndX</param>
		/// <param name="endY">Single EndY</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddConnector(NetOffice.OfficeApi.Enums.MsoConnectorType type, Single beginX, Single beginY, Single endX, Single endY)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, beginX, beginY, endX, endY);
			object returnItem = Invoker.MethodReturn(this, "AddConnector", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="safeArrayOfPoints">object SafeArrayOfPoints</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddCurve(object safeArrayOfPoints)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(safeArrayOfPoints);
			object returnItem = Invoker.MethodReturn(this, "AddCurve", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation Orientation</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddLabel(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(orientation, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddLabel", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="beginX">Single BeginX</param>
		/// <param name="beginY">Single BeginY</param>
		/// <param name="endX">Single EndX</param>
		/// <param name="endY">Single EndY</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddLine(Single beginX, Single beginY, Single endX, Single endY)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(beginX, beginY, endX, endY);
			object returnItem = Invoker.MethodReturn(this, "AddLine", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState LinkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddPicture(string fileName, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, linkToFile, saveWithDocument, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddPicture", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState LinkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddPicture(string fileName, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, linkToFile, saveWithDocument, left, top);
			object returnItem = Invoker.MethodReturn(this, "AddPicture", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState LinkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddPicture(string fileName, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, linkToFile, saveWithDocument, left, top, width);
			object returnItem = Invoker.MethodReturn(this, "AddPicture", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="safeArrayOfPoints">object SafeArrayOfPoints</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddPolyline(object safeArrayOfPoints)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(safeArrayOfPoints);
			object returnItem = Invoker.MethodReturn(this, "AddPolyline", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoAutoShapeType Type</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddShape(NetOffice.OfficeApi.Enums.MsoAutoShapeType type, Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddShape", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="presetTextEffect">NetOffice.OfficeApi.Enums.MsoPresetTextEffect PresetTextEffect</param>
		/// <param name="text">string Text</param>
		/// <param name="fontName">string FontName</param>
		/// <param name="fontSize">Single FontSize</param>
		/// <param name="fontBold">NetOffice.OfficeApi.Enums.MsoTriState FontBold</param>
		/// <param name="fontItalic">NetOffice.OfficeApi.Enums.MsoTriState FontItalic</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddTextEffect(NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect, string text, string fontName, Single fontSize, NetOffice.OfficeApi.Enums.MsoTriState fontBold, NetOffice.OfficeApi.Enums.MsoTriState fontItalic, Single left, Single top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(presetTextEffect, text, fontName, fontSize, fontBold, fontItalic, left, top);
			object returnItem = Invoker.MethodReturn(this, "AddTextEffect", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation Orientation</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddTextbox(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(orientation, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddTextbox", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType EditingType</param>
		/// <param name="x1">Single X1</param>
		/// <param name="y1">Single Y1</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.FreeformBuilder BuildFreeform(NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(editingType, x1, y1);
			object returnItem = Invoker.MethodReturn(this, "BuildFreeform", paramsArray);
			NetOffice.OfficeApi.FreeformBuilder newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.FreeformBuilder.LateBindingApiWrapperType) as NetOffice.OfficeApi.FreeformBuilder;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.ShapeRange Range(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "Range", paramsArray);
			NetOffice.OfficeApi.ShapeRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.ShapeRange.LateBindingApiWrapperType) as NetOffice.OfficeApi.ShapeRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public void SelectAll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SelectAll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoDiagramType Type</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddDiagram(NetOffice.OfficeApi.Enums.MsoDiagramType type, Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddDiagram", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddCanvas(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddCanvas", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddChart(object type, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddChart", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddChart()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddChart", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddChart(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.MethodReturn(this, "AddChart", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddChart(object type, object left)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, left);
			object returnItem = Invoker.MethodReturn(this, "AddChart", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddChart(object type, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, left, top);
			object returnItem = Invoker.MethodReturn(this, "AddChart", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddChart(object type, object left, object top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, left, top, width);
			object returnItem = Invoker.MethodReturn(this, "AddChart", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="numRows">Int32 NumRows</param>
		/// <param name="numColumns">Int32 NumColumns</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddTable(Int32 numRows, Int32 numColumns, Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numRows, numColumns, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddTable", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout Layout</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersionAttribute("Office", 14,15,16)]
		public NetOffice.OfficeApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(layout, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddSmartArt", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout Layout</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public NetOffice.OfficeApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(layout);
			object returnItem = Invoker.MethodReturn(this, "AddSmartArt", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout Layout</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public NetOffice.OfficeApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(layout, left);
			object returnItem = Invoker.MethodReturn(this, "AddSmartArt", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout Layout</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public NetOffice.OfficeApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(layout, left, top);
			object returnItem = Invoker.MethodReturn(this, "AddSmartArt", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// 
		/// </summary>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout Layout</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public NetOffice.OfficeApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(layout, left, top, width);
			object returnItem = Invoker.MethodReturn(this, "AddSmartArt", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// 
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		/// <param name="newLayout">optional bool NewLayout = true</param>
		[SupportByVersionAttribute("Office", 15, 16)]
		public NetOffice.OfficeApi.Shape AddChart2(object style, object type, object left, object top, object width, object height, object newLayout)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style, type, left, top, width, height, newLayout);
			object returnItem = Invoker.MethodReturn(this, "AddChart2", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 15, 16)]
		public NetOffice.OfficeApi.Shape AddChart2()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddChart2", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// 
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 15, 16)]
		public NetOffice.OfficeApi.Shape AddChart2(object style)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style);
			object returnItem = Invoker.MethodReturn(this, "AddChart2", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// 
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 15, 16)]
		public NetOffice.OfficeApi.Shape AddChart2(object style, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style, type);
			object returnItem = Invoker.MethodReturn(this, "AddChart2", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// 
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 15, 16)]
		public NetOffice.OfficeApi.Shape AddChart2(object style, object type, object left)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style, type, left);
			object returnItem = Invoker.MethodReturn(this, "AddChart2", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// 
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 15, 16)]
		public NetOffice.OfficeApi.Shape AddChart2(object style, object type, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style, type, left, top);
			object returnItem = Invoker.MethodReturn(this, "AddChart2", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// 
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 15, 16)]
		public NetOffice.OfficeApi.Shape AddChart2(object style, object type, object left, object top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style, type, left, top, width);
			object returnItem = Invoker.MethodReturn(this, "AddChart2", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// 
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 15, 16)]
		public NetOffice.OfficeApi.Shape AddChart2(object style, object type, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style, type, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddChart2", paramsArray);
			NetOffice.OfficeApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.Shape.LateBindingApiWrapperType) as NetOffice.OfficeApi.Shape;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.OfficeApi.Shape> Member
        
        /// <summary>
		/// SupportByVersionAttribute Office, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
       public IEnumerator<NetOffice.OfficeApi.Shape> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.OfficeApi.Shape item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Office, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}