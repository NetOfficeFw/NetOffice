using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface ShapeRange 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Publisher", 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("00021260-0000-0000-C000-000000000046")]
	public interface ShapeRange : ICOMObject, NetOffice.CollectionsGeneric.IEnumerableProvider<NetOffice.PublisherApi.Shape>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Adjustments Adjustments { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoAutoShapeType AutoShapeType { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.CalloutFormat Callout { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 ConnectionSiteCount { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState Connector { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ConnectorFormat ConnectorFormat { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.FillFormat Fill { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.GroupShapes GroupItems { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		object Height { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState HorizontalFlip { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		object Left { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.LineFormat Line { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState LockAspectRatio { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ShapeNodes Nodes { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Single Rotation { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.PictureFormat PictureFormat { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ShadowFormat Shadow { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextEffectFormat TextEffect { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextFrame TextFrame { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ThreeDFormat ThreeD { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		object Top { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Enums.PbShapeType Type { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState VerticalFlip { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		object Vertices { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		object Width { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 ZOrderPosition { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		string AlternativeText { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState HasTable { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState HasTextFrame { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Hyperlink Hyperlink { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Enums.PbWizardTag WizardTag { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 ID { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.LinkFormat LinkFormat { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.OLEFormat OLEFormat { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Table Table { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Tags Tags { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoBlackWhiteMode BlackWhiteMode { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 WizardTagInstance { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.WrapFormat TextWrap { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Wizard Wizard { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState IsInline { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Enums.PbInlineAlignment InlineAlignment { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.TextRange InlineTextRange { get; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		NetOffice.PublisherApi.SoftEdgeFormat SoftEdge { get; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		NetOffice.PublisherApi.GlowFormat Glow { get; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		NetOffice.PublisherApi.ReflectionFormat Reflection { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PublisherApi.Shape this[object index] { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void Apply();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void Cut();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void Copy();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ShapeRange Duplicate();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="flipCmd">NetOffice.OfficeApi.Enums.MsoFlipCmd flipCmd</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void Flip(NetOffice.OfficeApi.Enums.MsoFlipCmd flipCmd);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbUnitType unit</param>
		[SupportByVersion("Publisher", 14,15,16)]
		Single GetLeft(NetOffice.PublisherApi.Enums.PbUnitType unit);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbUnitType unit</param>
		[SupportByVersion("Publisher", 14,15,16)]
		Single GetTop(NetOffice.PublisherApi.Enums.PbUnitType unit);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbUnitType unit</param>
		[SupportByVersion("Publisher", 14,15,16)]
		Single GetHeight(NetOffice.PublisherApi.Enums.PbUnitType unit);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="unit">NetOffice.PublisherApi.Enums.PbUnitType unit</param>
		[SupportByVersion("Publisher", 14,15,16)]
		Single GetWidth(NetOffice.PublisherApi.Enums.PbUnitType unit);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="increment">object increment</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void IncrementLeft(object increment);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void IncrementRotation(Single increment);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="increment">object increment</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void IncrementTop(object increment);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void PickUp();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void RerouteConnections();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="factor">Single factor</param>
		/// <param name="relativeToOriginalSize">NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize</param>
		/// <param name="fScale">optional NetOffice.OfficeApi.Enums.MsoScaleFrom fScale</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void ScaleHeight(Single factor, NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize, object fScale);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="factor">Single factor</param>
		/// <param name="relativeToOriginalSize">NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ScaleHeight(Single factor, NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="factor">Single factor</param>
		/// <param name="relativeToOriginalSize">NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize</param>
		/// <param name="fScale">optional NetOffice.OfficeApi.Enums.MsoScaleFrom fScale</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void ScaleWidth(Single factor, NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize, object fScale);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="factor">Single factor</param>
		/// <param name="relativeToOriginalSize">NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void ScaleWidth(Single factor, NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="replace">optional object replace</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void Select(object replace);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void Select();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void SetShapesDefaultProperties();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ShapeRange Ungroup();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="zOrderCmd">NetOffice.OfficeApi.Enums.MsoZOrderCmd zOrderCmd</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void ZOrder(NetOffice.OfficeApi.Enums.MsoZOrderCmd zOrderCmd);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="alignCmd">NetOffice.OfficeApi.Enums.MsoAlignCmd alignCmd</param>
		/// <param name="relativeTo">NetOffice.OfficeApi.Enums.MsoTriState relativeTo</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void Align(NetOffice.OfficeApi.Enums.MsoAlignCmd alignCmd, NetOffice.OfficeApi.Enums.MsoTriState relativeTo);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="distributeCmd">NetOffice.OfficeApi.Enums.MsoDistributeCmd distributeCmd</param>
		/// <param name="relativeTo">NetOffice.OfficeApi.Enums.MsoTriState relativeTo</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void Distribute(NetOffice.OfficeApi.Enums.MsoDistributeCmd distributeCmd, NetOffice.OfficeApi.Enums.MsoTriState relativeTo);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape Group();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape Regroup();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.PublisherApi.TextRange range</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void MoveIntoTextFlow(NetOffice.PublisherApi.TextRange range);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void MoveOutOfTextFlow();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void AddToCatalogMergeArea();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void RemoveFromCatalogMergeArea();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="pbResolution">optional NetOffice.PublisherApi.Enums.PbPictureResolution pbResolution = 0</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void SaveAsPicture(string filename, object pbResolution);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		void SaveAsPicture(string filename);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.BuildingBlock SaveAsBuildingBlock(string name);

        #endregion

        #region IEnumerable<NetOffice.PublisherApi.Shape>

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        new IEnumerator<NetOffice.PublisherApi.Shape> GetEnumerator();

        #endregion
    }
}
