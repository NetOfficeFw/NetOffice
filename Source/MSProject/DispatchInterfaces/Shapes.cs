using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface Shapes 
	/// SupportByVersion MSProject, 11
	/// </summary>
	[SupportByVersion("MSProject", 11)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class Shapes : COMObject, IEnumerableProvider<NetOffice.MSProjectApi.Shape>
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
                    _type = typeof(Shapes);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Shapes(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		[SupportByVersion("MSProject", 11), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSProjectApi.Shape get_Value(object index)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Shape>(this, "Value", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Alias for get_Value
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11), Redirect("get_Value")]
		public NetOffice.MSProjectApi.Shape Value(object index)
		{
			return get_Value(index);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape Background
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Shape>(this, "Background", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape Default
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Shape>(this, "Default", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.MSProjectApi.Shape this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "Item", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoCalloutType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddCallout(NetOffice.OfficeApi.Enums.MsoCalloutType type, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddCallout", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoConnectorType type</param>
		/// <param name="beginX">Single beginX</param>
		/// <param name="beginY">Single beginY</param>
		/// <param name="endX">Single endX</param>
		/// <param name="endY">Single endY</param>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddConnector(NetOffice.OfficeApi.Enums.MsoConnectorType type, Single beginX, Single beginY, Single endX, Single endY)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddConnector", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, new object[]{ type, beginX, beginY, endX, endY });
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddCurve(object safeArrayOfPoints)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddCurve", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, safeArrayOfPoints);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation orientation</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddLabel(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddLabel", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, new object[]{ orientation, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="beginX">Single beginX</param>
		/// <param name="beginY">Single beginY</param>
		/// <param name="endX">Single endX</param>
		/// <param name="endY">Single endY</param>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddLine(Single beginX, Single beginY, Single endX, Single endY)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddLine", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, beginX, beginY, endX, endY);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddPicture", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, new object[]{ filename, linkToFile, saveWithDocument, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddPicture", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, new object[]{ filename, linkToFile, saveWithDocument, left, top });
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">optional Single Width = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddPicture", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, new object[]{ filename, linkToFile, saveWithDocument, left, top, width });
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddPolyline(object safeArrayOfPoints)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddPolyline", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, safeArrayOfPoints);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoAutoShapeType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddShape(NetOffice.OfficeApi.Enums.MsoAutoShapeType type, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddShape", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="presetTextEffect">NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect</param>
		/// <param name="text">string text</param>
		/// <param name="fontName">string fontName</param>
		/// <param name="fontSize">Single fontSize</param>
		/// <param name="fontBold">NetOffice.OfficeApi.Enums.MsoTriState fontBold</param>
		/// <param name="fontItalic">NetOffice.OfficeApi.Enums.MsoTriState fontItalic</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddTextEffect(NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect, string text, string fontName, Single fontSize, NetOffice.OfficeApi.Enums.MsoTriState fontBold, NetOffice.OfficeApi.Enums.MsoTriState fontItalic, Single left, Single top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddTextEffect", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, new object[]{ presetTextEffect, text, fontName, fontSize, fontBold, fontItalic, left, top });
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation orientation</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddTextbox(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddTextbox", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, new object[]{ orientation, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">Single x1</param>
		/// <param name="y1">Single y1</param>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.OfficeApi.FreeformBuilder BuildFreeform(NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.FreeformBuilder>(this, "BuildFreeform", NetOffice.OfficeApi.FreeformBuilder.LateBindingApiWrapperType, editingType, x1, y1);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.ShapeRange Range(object index)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.ShapeRange>(this, "Range", NetOffice.MSProjectApi.ShapeRange.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public void SelectAll()
		{
			 Factory.ExecuteMethod(this, "SelectAll");
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		/// <param name="newLayout">optional bool NewLayout = true</param>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddChart(object style, object type, object left, object top, object width, object height, object newLayout)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddChart", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, new object[]{ style, type, left, top, width, height, newLayout });
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddChart()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddChart", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddChart(object style)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddChart", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, style);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddChart(object style, object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddChart", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, style, type);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddChart(object style, object type, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddChart", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, style, type, left);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddChart(object style, object type, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddChart", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, style, type, left, top);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddChart(object style, object type, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddChart", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, new object[]{ style, type, left, top, width });
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddChart(object style, object type, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddChart", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, new object[]{ style, type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Shape AddTable(Int32 numRows, Int32 numColumns, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.Shape>(this, "AddTable", NetOffice.MSProjectApi.Shape.LateBindingApiWrapperType, new object[]{ numRows, numColumns, left, top, width, height });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.MSProjectApi.Shape>

        ICOMObject IEnumerableProvider<NetOffice.MSProjectApi.Shape>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.MSProjectApi.Shape>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.MSProjectApi.Shape>

        /// <summary>
        /// SupportByVersion MSProject, 11
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public IEnumerator<NetOffice.MSProjectApi.Shape> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.MSProjectApi.Shape item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion MSProject, 11
        /// </summary>
        [SupportByVersion("MSProject", 11)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}