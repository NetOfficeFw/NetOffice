using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface DiagramNode 
	/// SupportByVersion Excel, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
    [Duplicate("NetOffice.OfficeApi.DiagramNode")]
	public interface DiagramNode : NetOffice.OfficeApi._IMsoDispObj
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.DiagramNodeChildren Children { get; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Shape Shape { get; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.DiagramNode Root { get; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.IMsoDiagram Diagram { get; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoOrgChartLayoutType Layout { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Shape TextShape { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos = 2</param>
		/// <param name="nodeType">optional NetOffice.OfficeApi.Enums.MsoDiagramNodeType nodeType = 1</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.DiagramNode AddNode(object pos, object nodeType);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.DiagramNode AddNode();

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos = 2</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.DiagramNode AddNode(object pos);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pTargetNode">NetOffice.ExcelApi.DiagramNode pTargetNode</param>
		/// <param name="pos">NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void MoveNode(NetOffice.ExcelApi.DiagramNode pTargetNode, NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pTargetNode">NetOffice.ExcelApi.DiagramNode pTargetNode</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void ReplaceNode(NetOffice.ExcelApi.DiagramNode pTargetNode);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pTargetNode">NetOffice.ExcelApi.DiagramNode pTargetNode</param>
		/// <param name="swapChildren">optional bool swapChildren = true</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void SwapNode(NetOffice.ExcelApi.DiagramNode pTargetNode, object swapChildren);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pTargetNode">NetOffice.ExcelApi.DiagramNode pTargetNode</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void SwapNode(NetOffice.ExcelApi.DiagramNode pTargetNode);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="copyChildren">bool copyChildren</param>
		/// <param name="pTargetNode">NetOffice.ExcelApi.DiagramNode pTargetNode</param>
		/// <param name="pos">optional NetOffice.OfficeApi.Enums.MsoRelativeNodePosition pos = 2</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.DiagramNode CloneNode(bool copyChildren, NetOffice.ExcelApi.DiagramNode pTargetNode, object pos);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="copyChildren">bool copyChildren</param>
		/// <param name="pTargetNode">NetOffice.ExcelApi.DiagramNode pTargetNode</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.DiagramNode CloneNode(bool copyChildren, NetOffice.ExcelApi.DiagramNode pTargetNode);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="pReceivingNode">NetOffice.ExcelApi.DiagramNode pReceivingNode</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void TransferChildren(NetOffice.ExcelApi.DiagramNode pReceivingNode);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.DiagramNode NextNode();

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.DiagramNode PrevNode();

		#endregion
	}
}
