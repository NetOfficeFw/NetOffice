using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface DiagramNode 
	/// SupportByVersion Word, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface DiagramNode : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.DiagramNodeChildren Children { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Shape Shape { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.DiagramNode Root { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Diagram Diagram { get; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoOrgChartLayoutType Layout { get; set; }

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.Shape TextShape { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoDiagramNodeType NodeType = 1</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.DiagramNode AddNode(object pos, object nodeType);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.DiagramNode AddNode();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.DiagramNode AddNode(object pos);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="targetNode">NetOffice.WordApi.DiagramNode targetNode</param>
		/// <param name="pos">NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void MoveNode(out NetOffice.WordApi.DiagramNode targetNode, NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="targetNode">NetOffice.WordApi.DiagramNode targetNode</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void ReplaceNode(out NetOffice.WordApi.DiagramNode targetNode);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="targetNode">NetOffice.WordApi.DiagramNode targetNode</param>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = -1</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SwapNode(out NetOffice.WordApi.DiagramNode targetNode, object pos);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="targetNode">NetOffice.WordApi.DiagramNode targetNode</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void SwapNode(out NetOffice.WordApi.DiagramNode targetNode);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="copyChildren">bool copyChildren</param>
		/// <param name="targetNode">optional NetOffice.WordApi.DiagramNode TargetNode = 0</param>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition Pos = 2</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.DiagramNode CloneNode(bool copyChildren, object targetNode, object pos);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="copyChildren">bool copyChildren</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.DiagramNode CloneNode(bool copyChildren);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="copyChildren">bool copyChildren</param>
		/// <param name="targetNode">optional NetOffice.WordApi.DiagramNode TargetNode = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.DiagramNode CloneNode(bool copyChildren, object targetNode);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="receivingNode">NetOffice.WordApi.DiagramNode receivingNode</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		void TransferChildren(out NetOffice.WordApi.DiagramNode receivingNode);

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.DiagramNode NextNode();

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		NetOffice.WordApi.DiagramNode PrevNode();

		#endregion
	}
}
