using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.MSProjectApi
{
	///<summary>
	/// DispatchInterface Shapes 
	/// SupportByVersion MSProject, 11
	///</summary>
	[SupportByVersionAttribute("MSProject", 11)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Shapes : COMObject ,IEnumerable<NetOffice.MSProjectApi.Shape>
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
		/// SupportByVersion MSProject 11
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11)]
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
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11)]
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
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("MSProject", 11)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSProjectApi.Shape get_Value(object index)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Value", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Alias for get_Value
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape Value(object index)
		{
			return get_Value(index);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape Background
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Background", paramsArray);
				NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape Default
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Default", paramsArray);
				NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("MSProject", 11)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.MSProjectApi.Shape this[object index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoCalloutType Type</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddCallout(NetOffice.OfficeApi.Enums.MsoCalloutType type, Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddCallout", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoConnectorType Type</param>
		/// <param name="beginX">Single BeginX</param>
		/// <param name="beginY">Single BeginY</param>
		/// <param name="endX">Single EndX</param>
		/// <param name="endY">Single EndY</param>
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddConnector(NetOffice.OfficeApi.Enums.MsoConnectorType type, Single beginX, Single beginY, Single endX, Single endY)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, beginX, beginY, endX, endY);
			object returnItem = Invoker.MethodReturn(this, "AddConnector", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="safeArrayOfPoints">object SafeArrayOfPoints</param>
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddCurve(object safeArrayOfPoints)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(safeArrayOfPoints);
			object returnItem = Invoker.MethodReturn(this, "AddCurve", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation Orientation</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddLabel(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(orientation, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddLabel", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="beginX">Single BeginX</param>
		/// <param name="beginY">Single BeginY</param>
		/// <param name="endX">Single EndX</param>
		/// <param name="endY">Single EndY</param>
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddLine(Single beginX, Single beginY, Single endX, Single endY)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(beginX, beginY, endX, endY);
			object returnItem = Invoker.MethodReturn(this, "AddLine", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState LinkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, linkToFile, saveWithDocument, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddPicture", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState LinkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, linkToFile, saveWithDocument, left, top);
			object returnItem = Invoker.MethodReturn(this, "AddPicture", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState LinkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">optional Single Width = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, linkToFile, saveWithDocument, left, top, width);
			object returnItem = Invoker.MethodReturn(this, "AddPicture", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="safeArrayOfPoints">object SafeArrayOfPoints</param>
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddPolyline(object safeArrayOfPoints)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(safeArrayOfPoints);
			object returnItem = Invoker.MethodReturn(this, "AddPolyline", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoAutoShapeType Type</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddShape(NetOffice.OfficeApi.Enums.MsoAutoShapeType type, Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddShape", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
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
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddTextEffect(NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect, string text, string fontName, Single fontSize, NetOffice.OfficeApi.Enums.MsoTriState fontBold, NetOffice.OfficeApi.Enums.MsoTriState fontItalic, Single left, Single top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(presetTextEffect, text, fontName, fontSize, fontBold, fontItalic, left, top);
			object returnItem = Invoker.MethodReturn(this, "AddTextEffect", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation Orientation</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddTextbox(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(orientation, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddTextbox", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType EditingType</param>
		/// <param name="x1">Single X1</param>
		/// <param name="y1">Single Y1</param>
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.OfficeApi.FreeformBuilder BuildFreeform(NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(editingType, x1, y1);
			object returnItem = Invoker.MethodReturn(this, "BuildFreeform", paramsArray);
			NetOffice.OfficeApi.FreeformBuilder newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.FreeformBuilder.LateBindingApiWrapperType) as NetOffice.OfficeApi.FreeformBuilder;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.ShapeRange Range(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "Range", paramsArray);
			NetOffice.MSProjectApi.ShapeRange newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.ShapeRange.LateBindingApiWrapperType) as NetOffice.MSProjectApi.ShapeRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11)]
		public void SelectAll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SelectAll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		/// <param name="newLayout">optional bool NewLayout = true</param>
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddChart(object style, object type, object left, object top, object width, object height, object newLayout)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style, type, left, top, width, height, newLayout);
			object returnItem = Invoker.MethodReturn(this, "AddChart", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddChart()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddChart", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddChart(object style)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style);
			object returnItem = Invoker.MethodReturn(this, "AddChart", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddChart(object style, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style, type);
			object returnItem = Invoker.MethodReturn(this, "AddChart", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddChart(object style, object type, object left)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style, type, left);
			object returnItem = Invoker.MethodReturn(this, "AddChart", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddChart(object style, object type, object left, object top)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style, type, left, top);
			object returnItem = Invoker.MethodReturn(this, "AddChart", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddChart(object style, object type, object left, object top, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style, type, left, top, width);
			object returnItem = Invoker.MethodReturn(this, "AddChart", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddChart(object style, object type, object left, object top, object width, object height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style, type, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddChart", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="numRows">Int32 NumRows</param>
		/// <param name="numColumns">Int32 NumColumns</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		[SupportByVersionAttribute("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddTable(Int32 numRows, Int32 numColumns, Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numRows, numColumns, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddTable", paramsArray);
			NetOffice.MSProjectApi.Shape newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Shape;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.MSProjectApi.Shape> Member
        
        /// <summary>
		/// SupportByVersionAttribute MSProject, 11
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11)]
       public IEnumerator<NetOffice.MSProjectApi.Shape> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.MSProjectApi.Shape item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute MSProject, 11
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}