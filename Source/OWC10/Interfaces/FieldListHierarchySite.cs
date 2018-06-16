using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// Interface FieldListHierarchySite 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("FA99DB40-2043-11D3-854E-00C04FAC67D7")]
	public interface FieldListHierarchySite : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="nOldNodeId">Int32 nOldNodeId</param>
		/// <param name="nOldTypeId">Int32 nOldTypeId</param>
		/// <param name="pfPrevent">Int32 pfPrevent</param>
		[SupportByVersion("OWC10", 1)]
		Int32 PreSelect(Int32 nNodeId, Int32 nTypeId, Int32 nOldNodeId, Int32 nOldTypeId, out Int32 pfPrevent);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="nOldNodeId">Int32 nOldNodeId</param>
		/// <param name="nOldTypeId">Int32 nOldTypeId</param>
		[SupportByVersion("OWC10", 1)]
		Int32 PostSelect(Int32 nNodeId, Int32 nTypeId, Int32 nOldNodeId, Int32 nOldTypeId);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="fExpand">Int32 fExpand</param>
		/// <param name="pfPrevent">Int32 pfPrevent</param>
		[SupportByVersion("OWC10", 1)]
		Int32 PreExpand(Int32 nNodeId, Int32 nTypeId, Int32 fExpand, out Int32 pfPrevent);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="fExpand">Int32 fExpand</param>
		[SupportByVersion("OWC10", 1)]
		Int32 PostExpand(Int32 nNodeId, Int32 nTypeId, Int32 fExpand);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="ppobject">object ppobject</param>
		/// <param name="ppPivotView">object ppPivotView</param>
		/// <param name="pfPrevent">Int32 pfPrevent</param>
		[SupportByVersion("OWC10", 1)]
		Int32 PreDrag(Int32 nNodeId, Int32 nTypeId, out object ppobject, out object ppPivotView, out Int32 pfPrevent);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="hRes">Int32 hRes</param>
		[SupportByVersion("OWC10", 1)]
		Int32 PostDrag(Int32 nNodeId, Int32 nTypeId, Int32 hRes);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		[SupportByVersion("OWC10", 1)]
		Int32 PopulateChildren(Int32 nNodeId, Int32 nTypeId);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="hMenu">UIntPtr hMenu</param>
		/// <param name="pfPrevent">Int32 pfPrevent</param>
		[SupportByVersion("OWC10", 1)]
		Int32 ContextMenu(Int32 nNodeId, Int32 nTypeId, UIntPtr hMenu, out Int32 pfPrevent);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="wid">UIntPtr wid</param>
		[SupportByVersion("OWC10", 1)]
		Int32 DoCommand(Int32 nNodeId, Int32 nTypeId, UIntPtr wid);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		[SupportByVersion("OWC10", 1)]
		Int32 DoubleClick(Int32 nNodeId, Int32 nTypeId);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		[SupportByVersion("OWC10", 1)]
		Int32 PostDelete(Int32 nNodeId, Int32 nTypeId);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nSelMask">Int32 nSelMask</param>
		[SupportByVersion("OWC10", 1)]
		Int32 PostMSelect(Int32 nSelMask);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		[SupportByVersion("OWC10", 1)]
		Int32 Click(Int32 nNodeId, Int32 nTypeId);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="nMsg">Int32 nMsg</param>
		/// <param name="nwParam">Int32 nwParam</param>
		/// <param name="nlParam">Int32 nlParam</param>
		/// <param name="pfStopProcessing">Int32 pfStopProcessing</param>
		[SupportByVersion("OWC10", 1)]
		Int32 KeyEvent(Int32 nNodeId, Int32 nTypeId, Int32 nMsg, Int32 nwParam, Int32 nlParam, out Int32 pfStopProcessing);

		#endregion
	}
}
