using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface IShapes 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "_Default")]
	public class IShapes : COMObject, IEnumerableProvider<NetOffice.ExcelApi.Shape>
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
                    _type = typeof(IShapes);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IShapes(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IShapes(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IShapes(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IShapes(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IShapes(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IShapes(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IShapes() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IShapes(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.ExcelApi.ShapeRange get_Range(object index)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ShapeRange>(this, "Range", NetOffice.ExcelApi.ShapeRange.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Range
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_Range")]
		public NetOffice.ExcelApi.ShapeRange Range(object index)
		{
			return get_Range(index);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.ExcelApi.Shape this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "_Default", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoCalloutType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddCallout(NetOffice.OfficeApi.Enums.MsoCalloutType type, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddCallout", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoConnectorType type</param>
		/// <param name="beginX">Single beginX</param>
		/// <param name="beginY">Single beginY</param>
		/// <param name="endX">Single endX</param>
		/// <param name="endY">Single endY</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddConnector(NetOffice.OfficeApi.Enums.MsoConnectorType type, Single beginX, Single beginY, Single endX, Single endY)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddConnector", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ type, beginX, beginY, endX, endY });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddCurve(object safeArrayOfPoints)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddCurve", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, safeArrayOfPoints);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation orientation</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddLabel(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddLabel", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ orientation, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="beginX">Single beginX</param>
		/// <param name="beginY">Single beginY</param>
		/// <param name="endX">Single endX</param>
		/// <param name="endY">Single endY</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddLine(Single beginX, Single beginY, Single endX, Single endY)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddLine", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, beginX, beginY, endX, endY);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddPicture", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ filename, linkToFile, saveWithDocument, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddPolyline(object safeArrayOfPoints)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddPolyline", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, safeArrayOfPoints);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoAutoShapeType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddShape(NetOffice.OfficeApi.Enums.MsoAutoShapeType type, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddShape", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="presetTextEffect">NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect</param>
		/// <param name="text">string text</param>
		/// <param name="fontName">string fontName</param>
		/// <param name="fontSize">Single fontSize</param>
		/// <param name="fontBold">NetOffice.OfficeApi.Enums.MsoTriState fontBold</param>
		/// <param name="fontItalic">NetOffice.OfficeApi.Enums.MsoTriState fontItalic</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddTextEffect(NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect, string text, string fontName, Single fontSize, NetOffice.OfficeApi.Enums.MsoTriState fontBold, NetOffice.OfficeApi.Enums.MsoTriState fontItalic, Single left, Single top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddTextEffect", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ presetTextEffect, text, fontName, fontSize, fontBold, fontItalic, left, top });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation orientation</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddTextbox(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddTextbox", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ orientation, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">Single x1</param>
		/// <param name="y1">Single y1</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.FreeformBuilder BuildFreeform(NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.FreeformBuilder>(this, "BuildFreeform", NetOffice.ExcelApi.FreeformBuilder.LateBindingApiWrapperType, editingType, x1, y1);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 SelectAll()
		{
			return Factory.ExecuteInt32MethodGet(this, "SelectAll");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFormControl type</param>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddFormControl(NetOffice.ExcelApi.Enums.XlFormControl type, Int32 left, Int32 top, Int32 width, Int32 height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddFormControl", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="classType">optional object classType</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddOLEObject(object classType, object filename, object link, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddOLEObject", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ classType, filename, link, displayAsIcon, iconFileName, iconIndex, iconLabel, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddOLEObject()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddOLEObject", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="classType">optional object classType</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddOLEObject(object classType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddOLEObject", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, classType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="classType">optional object classType</param>
		/// <param name="filename">optional object filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddOLEObject(object classType, object filename)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddOLEObject", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, classType, filename);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="classType">optional object classType</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="link">optional object link</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddOLEObject(object classType, object filename, object link)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddOLEObject", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, classType, filename, link);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="classType">optional object classType</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddOLEObject(object classType, object filename, object link, object displayAsIcon)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddOLEObject", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, classType, filename, link, displayAsIcon);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="classType">optional object classType</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddOLEObject(object classType, object filename, object link, object displayAsIcon, object iconFileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddOLEObject", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ classType, filename, link, displayAsIcon, iconFileName });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="classType">optional object classType</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddOLEObject(object classType, object filename, object link, object displayAsIcon, object iconFileName, object iconIndex)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddOLEObject", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ classType, filename, link, displayAsIcon, iconFileName, iconIndex });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="classType">optional object classType</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddOLEObject(object classType, object filename, object link, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddOLEObject", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ classType, filename, link, displayAsIcon, iconFileName, iconIndex, iconLabel });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="classType">optional object classType</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddOLEObject(object classType, object filename, object link, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddOLEObject", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ classType, filename, link, displayAsIcon, iconFileName, iconIndex, iconLabel, left });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="classType">optional object classType</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddOLEObject(object classType, object filename, object link, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddOLEObject", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ classType, filename, link, displayAsIcon, iconFileName, iconIndex, iconLabel, left, top });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="classType">optional object classType</param>
		/// <param name="filename">optional object filename</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddOLEObject(object classType, object filename, object link, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddOLEObject", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ classType, filename, link, displayAsIcon, iconFileName, iconIndex, iconLabel, left, top, width });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoDiagramType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddDiagram(NetOffice.OfficeApi.Enums.MsoDiagramType type, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddDiagram", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddCanvas(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddCanvas", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="xlChartType">optional object xlChartType</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddChart(object xlChartType, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddChart", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ xlChartType, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddChart()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddChart", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="xlChartType">optional object xlChartType</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddChart(object xlChartType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddChart", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, xlChartType);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="xlChartType">optional object xlChartType</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddChart(object xlChartType, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddChart", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, xlChartType, left);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="xlChartType">optional object xlChartType</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddChart(object xlChartType, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddChart", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, xlChartType, left, top);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="xlChartType">optional object xlChartType</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Shape AddChart(object xlChartType, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddChart", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, xlChartType, left, top, width);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddSmartArt", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ layout, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddSmartArt", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, layout);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddSmartArt", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, layout, left);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddSmartArt", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, layout, left, top);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddSmartArt", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, layout, left, top, width);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="style">optional object style</param>
		/// <param name="xlChartType">optional object xlChartType</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		/// <param name="newLayout">optional object newLayout</param>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Shape AddChart2(object style, object xlChartType, object left, object top, object width, object height, object newLayout)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddChart2", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ style, xlChartType, left, top, width, height, newLayout });
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Shape AddChart2()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddChart2", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="style">optional object style</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Shape AddChart2(object style)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddChart2", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, style);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="style">optional object style</param>
		/// <param name="xlChartType">optional object xlChartType</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Shape AddChart2(object style, object xlChartType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddChart2", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, style, xlChartType);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="style">optional object style</param>
		/// <param name="xlChartType">optional object xlChartType</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Shape AddChart2(object style, object xlChartType, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddChart2", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, style, xlChartType, left);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="style">optional object style</param>
		/// <param name="xlChartType">optional object xlChartType</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Shape AddChart2(object style, object xlChartType, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddChart2", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, style, xlChartType, left, top);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="style">optional object style</param>
		/// <param name="xlChartType">optional object xlChartType</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Shape AddChart2(object style, object xlChartType, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddChart2", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ style, xlChartType, left, top, width });
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="style">optional object style</param>
		/// <param name="xlChartType">optional object xlChartType</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Shape AddChart2(object style, object xlChartType, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Shape>(this, "AddChart2", NetOffice.ExcelApi.Shape.LateBindingApiWrapperType, new object[]{ style, xlChartType, left, top, width, height });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.Shape>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.Shape>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.Shape>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.Shape>

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.ExcelApi.Shape> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.Shape item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}