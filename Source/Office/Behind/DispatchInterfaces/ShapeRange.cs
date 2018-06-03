using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface ShapeRange 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
    public class ShapeRange : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.ShapeRange
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
                    _type = typeof(ShapeRange);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ShapeRange() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Count
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Adjustments Adjustments
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Adjustments>(this, "Adjustments", typeof(NetOffice.OfficeApi.Adjustments));
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoAutoShapeType AutoShapeType
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoAutoShapeType>(this, "AutoShapeType");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "AutoShapeType", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoBlackWhiteMode BlackWhiteMode
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoBlackWhiteMode>(this, "BlackWhiteMode");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "BlackWhiteMode", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.CalloutFormat Callout
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CalloutFormat>(this, "Callout", typeof(NetOffice.OfficeApi.CalloutFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 ConnectionSiteCount
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "ConnectionSiteCount");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoTriState Connector
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Connector");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.ConnectorFormat ConnectorFormat
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ConnectorFormat>(this, "ConnectorFormat", typeof(NetOffice.OfficeApi.ConnectorFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.FillFormat Fill
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FillFormat>(this, "Fill", typeof(NetOffice.OfficeApi.FillFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.GroupShapes GroupItems
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.GroupShapes>(this, "GroupItems", typeof(NetOffice.OfficeApi.GroupShapes));
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public Single Height
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "Height");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Height", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoTriState HorizontalFlip
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "HorizontalFlip");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public Single Left
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "Left");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Left", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.LineFormat Line
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.LineFormat>(this, "Line", typeof(NetOffice.OfficeApi.LineFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoTriState LockAspectRatio
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "LockAspectRatio");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "LockAspectRatio", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
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
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.ShapeNodes Nodes
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ShapeNodes>(this, "Nodes", typeof(NetOffice.OfficeApi.ShapeNodes));
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public Single Rotation
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "Rotation");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Rotation", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.PictureFormat PictureFormat
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PictureFormat>(this, "PictureFormat", typeof(NetOffice.OfficeApi.PictureFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.ShadowFormat Shadow
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ShadowFormat>(this, "Shadow", typeof(NetOffice.OfficeApi.ShadowFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.TextEffectFormat TextEffect
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextEffectFormat>(this, "TextEffect", typeof(NetOffice.OfficeApi.TextEffectFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.TextFrame TextFrame
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextFrame>(this, "TextFrame", typeof(NetOffice.OfficeApi.TextFrame));
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.ThreeDFormat ThreeD
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ThreeDFormat>(this, "ThreeD", typeof(NetOffice.OfficeApi.ThreeDFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public Single Top
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "Top");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Top", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoShapeType Type
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoShapeType>(this, "Type");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoTriState VerticalFlip
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "VerticalFlip");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public object Vertices
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "Vertices");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoTriState Visible
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Visible");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "Visible", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public Single Width
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "Width");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Width", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 ZOrderPosition
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "ZOrderPosition");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Script Script
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Script>(this, "Script", typeof(NetOffice.OfficeApi.Script));
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public string AlternativeText
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "AlternativeText");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "AlternativeText", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoTriState HasDiagram
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "HasDiagram");
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.IMsoDiagram Diagram
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoDiagram>(this, "Diagram", typeof(NetOffice.OfficeApi.IMsoDiagram));
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoTriState HasDiagramNode
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "HasDiagramNode");
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.DiagramNode DiagramNode
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DiagramNode>(this, "DiagramNode", typeof(NetOffice.OfficeApi.DiagramNode));
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoTriState Child
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Child");
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Shape ParentGroup
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Shape>(this, "ParentGroup", typeof(NetOffice.OfficeApi.Shape));
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.OfficeApi.CanvasShapes CanvasItems
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CanvasShapes>(this, "CanvasItems", typeof(NetOffice.OfficeApi.CanvasShapes));
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public Int32 Id
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Id");
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public string RTF
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "RTF");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "RTF", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.TextFrame2 TextFrame2
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextFrame2>(this, "TextFrame2", typeof(NetOffice.OfficeApi.TextFrame2));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoTriState HasChart
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "HasChart");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.IMsoChart Chart
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChart>(this, "Chart", typeof(NetOffice.OfficeApi.IMsoChart));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoShapeStyleIndex ShapeStyle
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoShapeStyleIndex>(this, "ShapeStyle");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "ShapeStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoBackgroundStyleIndex BackgroundStyle
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoBackgroundStyleIndex>(this, "BackgroundStyle");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "BackgroundStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.SoftEdgeFormat SoftEdge
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SoftEdgeFormat>(this, "SoftEdge", typeof(NetOffice.OfficeApi.SoftEdgeFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.GlowFormat Glow
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.GlowFormat>(this, "Glow", typeof(NetOffice.OfficeApi.GlowFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.ReflectionFormat Reflection
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ReflectionFormat>(this, "Reflection", typeof(NetOffice.OfficeApi.ReflectionFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public string Title
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Title");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Title", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public NetOffice.OfficeApi.Shape this[object index]
        {
            get
            {
                return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "Item", typeof(NetOffice.OfficeApi.Shape), index);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="alignCmd">NetOffice.OfficeApi.Enums.MsoAlignCmd alignCmd</param>
        /// <param name="relativeTo">NetOffice.OfficeApi.Enums.MsoTriState relativeTo</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void Align(NetOffice.OfficeApi.Enums.MsoAlignCmd alignCmd, NetOffice.OfficeApi.Enums.MsoTriState relativeTo)
        {
            Factory.ExecuteMethod(this, "Align", alignCmd, relativeTo);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void Apply()
        {
            Factory.ExecuteMethod(this, "Apply");
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void Delete()
        {
            Factory.ExecuteMethod(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="distributeCmd">NetOffice.OfficeApi.Enums.MsoDistributeCmd distributeCmd</param>
        /// <param name="relativeTo">NetOffice.OfficeApi.Enums.MsoTriState relativeTo</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void Distribute(NetOffice.OfficeApi.Enums.MsoDistributeCmd distributeCmd, NetOffice.OfficeApi.Enums.MsoTriState relativeTo)
        {
            Factory.ExecuteMethod(this, "Distribute", distributeCmd, relativeTo);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.ShapeRange Duplicate()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.ShapeRange>(this, "Duplicate", typeof(NetOffice.OfficeApi.ShapeRange));
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="flipCmd">NetOffice.OfficeApi.Enums.MsoFlipCmd flipCmd</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void Flip(NetOffice.OfficeApi.Enums.MsoFlipCmd flipCmd)
        {
            Factory.ExecuteMethod(this, "Flip", flipCmd);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="increment">Single increment</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void IncrementLeft(Single increment)
        {
            Factory.ExecuteMethod(this, "IncrementLeft", increment);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="increment">Single increment</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void IncrementRotation(Single increment)
        {
            Factory.ExecuteMethod(this, "IncrementRotation", increment);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="increment">Single increment</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void IncrementTop(Single increment)
        {
            Factory.ExecuteMethod(this, "IncrementTop", increment);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Shape Group()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "Group", typeof(NetOffice.OfficeApi.Shape));
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void PickUp()
        {
            Factory.ExecuteMethod(this, "PickUp");
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Shape Regroup()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "Regroup", typeof(NetOffice.OfficeApi.Shape));
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void RerouteConnections()
        {
            Factory.ExecuteMethod(this, "RerouteConnections");
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="factor">Single factor</param>
        /// <param name="relativeToOriginalSize">NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize</param>
        /// <param name="fScale">optional NetOffice.OfficeApi.Enums.MsoScaleFrom fScale = 0</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void ScaleHeight(Single factor, NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize, object fScale)
        {
            Factory.ExecuteMethod(this, "ScaleHeight", factor, relativeToOriginalSize, fScale);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="factor">Single factor</param>
        /// <param name="relativeToOriginalSize">NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void ScaleHeight(Single factor, NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize)
        {
            Factory.ExecuteMethod(this, "ScaleHeight", factor, relativeToOriginalSize);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="factor">Single factor</param>
        /// <param name="relativeToOriginalSize">NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize</param>
        /// <param name="fScale">optional NetOffice.OfficeApi.Enums.MsoScaleFrom fScale = 0</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void ScaleWidth(Single factor, NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize, object fScale)
        {
            Factory.ExecuteMethod(this, "ScaleWidth", factor, relativeToOriginalSize, fScale);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="factor">Single factor</param>
        /// <param name="relativeToOriginalSize">NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void ScaleWidth(Single factor, NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize)
        {
            Factory.ExecuteMethod(this, "ScaleWidth", factor, relativeToOriginalSize);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="replace">optional object replace</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void Select(object replace)
        {
            Factory.ExecuteMethod(this, "Select", replace);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void Select()
        {
            Factory.ExecuteMethod(this, "Select");
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void SetShapesDefaultProperties()
        {
            Factory.ExecuteMethod(this, "SetShapesDefaultProperties");
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.ShapeRange Ungroup()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.ShapeRange>(this, "Ungroup", typeof(NetOffice.OfficeApi.ShapeRange));
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="zOrderCmd">NetOffice.OfficeApi.Enums.MsoZOrderCmd zOrderCmd</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void ZOrder(NetOffice.OfficeApi.Enums.MsoZOrderCmd zOrderCmd)
        {
            Factory.ExecuteMethod(this, "ZOrder", zOrderCmd);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="increment">Single increment</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public void CanvasCropLeft(Single increment)
        {
            Factory.ExecuteMethod(this, "CanvasCropLeft", increment);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="increment">Single increment</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public void CanvasCropTop(Single increment)
        {
            Factory.ExecuteMethod(this, "CanvasCropTop", increment);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="increment">Single increment</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public void CanvasCropRight(Single increment)
        {
            Factory.ExecuteMethod(this, "CanvasCropRight", increment);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="increment">Single increment</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public void CanvasCropBottom(Single increment)
        {
            Factory.ExecuteMethod(this, "CanvasCropBottom", increment);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public void Cut()
        {
            Factory.ExecuteMethod(this, "Cut");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public void Copy()
        {
            Factory.ExecuteMethod(this, "Copy");
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        /// <param name="mergeCmd">NetOffice.OfficeApi.Enums.MsoMergeCmd mergeCmd</param>
        /// <param name="primaryShape">optional NetOffice.OfficeApi.Shape PrimaryShape = 0</param>
        [SupportByVersion("Office", 15, 16)]
        public void MergeShapes(NetOffice.OfficeApi.Enums.MsoMergeCmd mergeCmd, object primaryShape)
        {
            Factory.ExecuteMethod(this, "MergeShapes", mergeCmd, primaryShape);
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        /// <param name="mergeCmd">NetOffice.OfficeApi.Enums.MsoMergeCmd mergeCmd</param>
        [CustomMethod]
        [SupportByVersion("Office", 15, 16)]
        public void MergeShapes(NetOffice.OfficeApi.Enums.MsoMergeCmd mergeCmd)
        {
            Factory.ExecuteMethod(this, "MergeShapes", mergeCmd);
        }

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.Shape>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.Shape>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.Shape>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.Shape>

        /// <summary>
        /// SupportByVersion Office, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.Shape> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.Shape item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}
