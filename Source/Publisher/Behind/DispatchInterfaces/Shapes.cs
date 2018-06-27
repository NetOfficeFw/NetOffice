using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.PublisherApi;

namespace NetOffice.PublisherApi.Behind
{
	/// <summary>
	/// DispatchInterface Shapes 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	public class Shapes : COMObject, NetOffice.PublisherApi.Shapes
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.PublisherApi.Shapes);
                return _contractType;
            }
        }
        private static Type _contractType;


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
                    _type = typeof(Shapes);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Shapes() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", typeof(NetOffice.PublisherApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.pbCanvasArrangementType CanvasArrangementType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.pbCanvasArrangementType>(this, "CanvasArrangementType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "CanvasArrangementType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 CanvasesCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CanvasesCount");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.PublisherApi.Shape this[object index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "Item", typeof(NetOffice.PublisherApi.Shape), index);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoCalloutType type</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddCallout(NetOffice.OfficeApi.Enums.MsoCalloutType type, object left, object top, object width, object height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddCallout", typeof(NetOffice.PublisherApi.Shape), new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoConnectorType type</param>
		/// <param name="beginX">object beginX</param>
		/// <param name="beginY">object beginY</param>
		/// <param name="endX">object endX</param>
		/// <param name="endY">object endY</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddConnector(NetOffice.OfficeApi.Enums.MsoConnectorType type, object beginX, object beginY, object endX, object endY)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddConnector", typeof(NetOffice.PublisherApi.Shape), new object[]{ type, beginX, beginY, endX, endY });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddCurve(object safeArrayOfPoints)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddCurve", typeof(NetOffice.PublisherApi.Shape), safeArrayOfPoints);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="orientation">NetOffice.PublisherApi.Enums.PbTextOrientation orientation</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddLabel(NetOffice.PublisherApi.Enums.PbTextOrientation orientation, object left, object top, object width, object height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddLabel", typeof(NetOffice.PublisherApi.Shape), new object[]{ orientation, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="beginX">object beginX</param>
		/// <param name="beginY">object beginY</param>
		/// <param name="endX">object endX</param>
		/// <param name="endY">object endY</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddLine(object beginX, object beginY, object endX, object endY)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddLine", typeof(NetOffice.PublisherApi.Shape), beginX, beginY, endX, endY);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		/// <param name="filename">optional string Filename = </param>
		/// <param name="link">optional NetOffice.OfficeApi.Enums.MsoTriState Link = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object filename, object link)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddOLEObject", typeof(NetOffice.PublisherApi.Shape), new object[]{ left, top, width, height, className, filename, link });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddOLEObject(object left, object top)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddOLEObject", typeof(NetOffice.PublisherApi.Shape), left, top);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddOLEObject", typeof(NetOffice.PublisherApi.Shape), left, top, width);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width, object height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddOLEObject", typeof(NetOffice.PublisherApi.Shape), left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width, object height, object className)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddOLEObject", typeof(NetOffice.PublisherApi.Shape), new object[]{ left, top, width, height, className });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		/// <param name="filename">optional string Filename = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object filename)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddOLEObject", typeof(NetOffice.PublisherApi.Shape), new object[]{ left, top, width, height, className, filename });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, object left, object top, object width, object height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddPicture", typeof(NetOffice.PublisherApi.Shape), new object[]{ filename, linkToFile, saveWithDocument, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, object left, object top)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddPicture", typeof(NetOffice.PublisherApi.Shape), new object[]{ filename, linkToFile, saveWithDocument, left, top });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, object left, object top, object width)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddPicture", typeof(NetOffice.PublisherApi.Shape), new object[]{ filename, linkToFile, saveWithDocument, left, top, width });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddPolyline(object safeArrayOfPoints)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddPolyline", typeof(NetOffice.PublisherApi.Shape), safeArrayOfPoints);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoAutoShapeType type</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddShape(NetOffice.OfficeApi.Enums.MsoAutoShapeType type, object left, object top, object width, object height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddShape", typeof(NetOffice.PublisherApi.Shape), new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		/// <param name="fixedSize">optional bool FixedSize = false</param>
		/// <param name="direction">optional NetOffice.PublisherApi.Enums.PbTableDirectionType Direction = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left, object top, object width, object height, object fixedSize, object direction)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddTable", typeof(NetOffice.PublisherApi.Shape), new object[]{ numRows, numColumns, left, top, width, height, fixedSize, direction });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left, object top, object width, object height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddTable", typeof(NetOffice.PublisherApi.Shape), new object[]{ numRows, numColumns, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		/// <param name="fixedSize">optional bool FixedSize = false</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left, object top, object width, object height, object fixedSize)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddTable", typeof(NetOffice.PublisherApi.Shape), new object[]{ numRows, numColumns, left, top, width, height, fixedSize });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="orientation">NetOffice.PublisherApi.Enums.PbTextOrientation orientation</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddTextbox(NetOffice.PublisherApi.Enums.PbTextOrientation orientation, object left, object top, object width, object height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddTextbox", typeof(NetOffice.PublisherApi.Shape), new object[]{ orientation, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="presetTextEffect">NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect</param>
		/// <param name="text">string text</param>
		/// <param name="fontName">string fontName</param>
		/// <param name="fontSize">object fontSize</param>
		/// <param name="fontBold">NetOffice.OfficeApi.Enums.MsoTriState fontBold</param>
		/// <param name="fontItalic">NetOffice.OfficeApi.Enums.MsoTriState fontItalic</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddTextEffect(NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect, string text, string fontName, object fontSize, NetOffice.OfficeApi.Enums.MsoTriState fontBold, NetOffice.OfficeApi.Enums.MsoTriState fontItalic, object left, object top)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddTextEffect", typeof(NetOffice.PublisherApi.Shape), new object[]{ presetTextEffect, text, fontName, fontSize, fontBold, fontItalic, left, top });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.PublisherApi.Enums.PbWebControlType type</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		/// <param name="launchPropertiesWindow">optional bool LaunchPropertiesWindow = false</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddWebControl(NetOffice.PublisherApi.Enums.PbWebControlType type, object left, object top, object width, object height, object launchPropertiesWindow)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddWebControl", typeof(NetOffice.PublisherApi.Shape), new object[]{ type, left, top, width, height, launchPropertiesWindow });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.PublisherApi.Enums.PbWebControlType type</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddWebControl(NetOffice.PublisherApi.Enums.PbWebControlType type, object left, object top, object width, object height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddWebControl", typeof(NetOffice.PublisherApi.Shape), new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">object x1</param>
		/// <param name="y1">object y1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.FreeformBuilder BuildFreeform(NetOffice.OfficeApi.Enums.MsoEditingType editingType, object x1, object y1)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.FreeformBuilder>(this, "BuildFreeform", typeof(NetOffice.PublisherApi.FreeformBuilder), editingType, x1, y1);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ShapeRange Paste()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ShapeRange>(this, "Paste", typeof(NetOffice.PublisherApi.ShapeRange));
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ShapeRange Range(object index)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ShapeRange>(this, "Range", typeof(NetOffice.PublisherApi.ShapeRange), index);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ShapeRange Range()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ShapeRange>(this, "Range", typeof(NetOffice.PublisherApi.ShapeRange));
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void SelectAll()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SelectAll");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizardGroup wizard</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		/// <param name="design">optional Int32 Design = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddGroupWizard(NetOffice.PublisherApi.Enums.PbWizardGroup wizard, object left, object top, object width, object height, object design)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddGroupWizard", typeof(NetOffice.PublisherApi.Shape), new object[]{ wizard, left, top, width, height, design });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizardGroup wizard</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddGroupWizard(NetOffice.PublisherApi.Enums.PbWizardGroup wizard, object left, object top)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddGroupWizard", typeof(NetOffice.PublisherApi.Shape), wizard, left, top);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizardGroup wizard</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddGroupWizard(NetOffice.PublisherApi.Enums.PbWizardGroup wizard, object left, object top, object width)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddGroupWizard", typeof(NetOffice.PublisherApi.Shape), wizard, left, top, width);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizardGroup wizard</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddGroupWizard(NetOffice.PublisherApi.Enums.PbWizardGroup wizard, object left, object top, object width, object height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddGroupWizard", typeof(NetOffice.PublisherApi.Shape), new object[]{ wizard, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizardTag">NetOffice.PublisherApi.Enums.PbWizardTag wizardTag</param>
		/// <param name="instance">optional Int32 Instance = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ShapeRange FindShapeByWizardTag(NetOffice.PublisherApi.Enums.PbWizardTag wizardTag, object instance)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ShapeRange>(this, "FindShapeByWizardTag", typeof(NetOffice.PublisherApi.ShapeRange), wizardTag, instance);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizardTag">NetOffice.PublisherApi.Enums.PbWizardTag wizardTag</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ShapeRange FindShapeByWizardTag(NetOffice.PublisherApi.Enums.PbWizardTag wizardTag)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ShapeRange>(this, "FindShapeByWizardTag", typeof(NetOffice.PublisherApi.ShapeRange), wizardTag);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object width</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddWebNavigationBar(string name, object left, object top, object width)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddWebNavigationBar", typeof(NetOffice.PublisherApi.Shape), name, left, top, width);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddWebNavigationBar(string name, object left, object top)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddWebNavigationBar", typeof(NetOffice.PublisherApi.Shape), name, left, top);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddCatalogMergeArea()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddCatalogMergeArea", typeof(NetOffice.PublisherApi.Shape));
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddEmptyPictureFrame(object left, object top, object width, object height)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddEmptyPictureFrame", typeof(NetOffice.PublisherApi.Shape), left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddEmptyPictureFrame(object left, object top)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddEmptyPictureFrame", typeof(NetOffice.PublisherApi.Shape), left, top);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddEmptyPictureFrame(object left, object top, object width)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddEmptyPictureFrame", typeof(NetOffice.PublisherApi.Shape), left, top, width);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="canvasId">Int32 canvasId</param>
		/// <param name="catalogMergeFieldType">NetOffice.PublisherApi.Enums.pbCatalogMergeFieldType catalogMergeFieldType</param>
		/// <param name="dbCol">Int32 dbCol</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void AddCatalogMergeFieldToCanvas(Int32 canvasId, NetOffice.PublisherApi.Enums.pbCatalogMergeFieldType catalogMergeFieldType, Int32 dbCol)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddCatalogMergeFieldToCanvas", canvasId, catalogMergeFieldType, dbCol);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="presetWordArt">NetOffice.PublisherApi.Enums.pbPresetWordArt presetWordArt</param>
		/// <param name="text">string text</param>
		/// <param name="fontName">string fontName</param>
		/// <param name="fontSize">object fontSize</param>
		/// <param name="fontBold">NetOffice.OfficeApi.Enums.MsoTriState fontBold</param>
		/// <param name="fontItalic">NetOffice.OfficeApi.Enums.MsoTriState fontItalic</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddWordArt(NetOffice.PublisherApi.Enums.pbPresetWordArt presetWordArt, string text, string fontName, object fontSize, NetOffice.OfficeApi.Enums.MsoTriState fontBold, NetOffice.OfficeApi.Enums.MsoTriState fontItalic, object left, object top)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddWordArt", typeof(NetOffice.PublisherApi.Shape), new object[]{ presetWordArt, text, fontName, fontSize, fontBold, fontItalic, left, top });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="bBlockIn">NetOffice.PublisherApi.BuildingBlock bBlockIn</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Shape AddBuildingBlock(NetOffice.PublisherApi.BuildingBlock bBlockIn, object left, object top)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddBuildingBlock", typeof(NetOffice.PublisherApi.Shape), bBlockIn, left, top);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.PublisherApi.Shape>

        ICOMObject IEnumerableProvider<NetOffice.PublisherApi.Shape>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.PublisherApi.Shape>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.PublisherApi.Shape>

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public virtual IEnumerator<NetOffice.PublisherApi.Shape> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.PublisherApi.Shape item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

