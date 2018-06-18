using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface Shape 
	/// SupportByVersion MSProject, 11
	/// </summary>
	[SupportByVersion("MSProject", 11)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("7300BAD0-C853-475E-80A7-2FA3A286F91F")]
	public interface Shape : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSProject", 11), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.Adjustments Adjustments { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.Enums.MsoAutoShapeType AutoShapeType { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.Enums.MsoBlackWhiteMode BlackWhiteMode { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.CalloutFormat Callout { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		Int32 ConnectionSiteCount { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.Enums.MsoTriState Connector { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.ConnectorFormat ConnectorFormat { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.FillFormat Fill { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.GroupShapes GroupItems { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		Single Height { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.Enums.MsoTriState HorizontalFlip { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		Single Left { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.LineFormat Line { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.Enums.MsoTriState LockAspectRatio { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.ShapeNodes Nodes { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		Single Rotation { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OfficeApi.PictureFormat PictureFormat { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.ShadowFormat Shadow { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.TextEffectFormat TextEffect { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.TextFrame TextFrame { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.ThreeDFormat ThreeD { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		Single Top { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.Enums.MsoShapeType Type { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.Enums.MsoTriState VerticalFlip { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		object Vertices { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.Enums.MsoTriState Visible { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		Single Width { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		Int32 ZOrderPosition { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OfficeApi.Script Script { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		string AlternativeText { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.Enums.MsoTriState Child { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape ParentGroup { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OfficeApi.CanvasShapes CanvasItems { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		Int32 ID { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string RTF { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.TextFrame2 TextFrame2 { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.Enums.MsoTriState HasChart { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Chart Chart { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.Enums.MsoShapeStyleIndex ShapeStyle { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.Enums.MsoBackgroundStyleIndex BackgroundStyle { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.SoftEdgeFormat SoftEdge { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.GlowFormat Glow { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.ReflectionFormat Reflection { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		string Title { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.Enums.MsoTriState HasTable { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.ReportTable Table { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		void Apply();

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		void Delete();

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape Duplicate();

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="flipCmd">NetOffice.OfficeApi.Enums.MsoFlipCmd flipCmd</param>
		[SupportByVersion("MSProject", 11)]
		void Flip(NetOffice.OfficeApi.Enums.MsoFlipCmd flipCmd);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("MSProject", 11)]
		void IncrementLeft(Single increment);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("MSProject", 11)]
		void IncrementRotation(Single increment);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="increment">Single increment</param>
		[SupportByVersion("MSProject", 11)]
		void IncrementTop(Single increment);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		void PickUp();

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		void RerouteConnections();

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="factor">Single factor</param>
		/// <param name="relativeToOriginalSize">NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize</param>
		/// <param name="fScale">optional NetOffice.OfficeApi.Enums.MsoScaleFrom fScale = 0</param>
		[SupportByVersion("MSProject", 11)]
		void ScaleHeight(Single factor, NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize, object fScale);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="factor">Single factor</param>
		/// <param name="relativeToOriginalSize">NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		void ScaleHeight(Single factor, NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="factor">Single factor</param>
		/// <param name="relativeToOriginalSize">NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize</param>
		/// <param name="fScale">optional NetOffice.OfficeApi.Enums.MsoScaleFrom fScale = 0</param>
		[SupportByVersion("MSProject", 11)]
		void ScaleWidth(Single factor, NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize, object fScale);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="factor">Single factor</param>
		/// <param name="relativeToOriginalSize">NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		void ScaleWidth(Single factor, NetOffice.OfficeApi.Enums.MsoTriState relativeToOriginalSize);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="replace">optional object replace</param>
		[SupportByVersion("MSProject", 11)]
		void Select(object replace);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		void Select();

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		void SetShapesDefaultProperties();

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.ShapeRange Ungroup();

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="zOrderCmd">NetOffice.OfficeApi.Enums.MsoZOrderCmd zOrderCmd</param>
		[SupportByVersion("MSProject", 11)]
		void ZOrder(NetOffice.OfficeApi.Enums.MsoZOrderCmd zOrderCmd);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="increment">Single increment</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSProject", 11)]
		void CanvasCropLeft(Single increment);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="increment">Single increment</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSProject", 11)]
		void CanvasCropTop(Single increment);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="increment">Single increment</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSProject", 11)]
		void CanvasCropRight(Single increment);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="increment">Single increment</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSProject", 11)]
		void CanvasCropBottom(Single increment);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		void Cut();

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		void Copy();

		#endregion
	}
}
