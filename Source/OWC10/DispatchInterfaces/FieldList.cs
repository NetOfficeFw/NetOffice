using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface FieldList 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("7BD1809E-0406-11D3-8549-00C04FAC67D7")]
    [CoClassSource(typeof(NetOffice.OWC10Api.FieldListControl))]
    public interface FieldList : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		Int32 ClipboardFormat { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		string InstanceID { get; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1), NativeResult]
		stdole.IFont Font { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		bool MultiSelect { get; set; }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.Enums.FieldListSelectRestriction SelectRestriction { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="bVisible">bool bVisible</param>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.FieldListHierarchy CreateHierarchy(bool bVisible);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="iWidth">Int32 iWidth</param>
		/// <param name="iHeight">Int32 iHeight</param>
		/// <param name="pip">stdole.IPicture pip</param>
		/// <param name="crMask">Int32 crMask</param>
		[SupportByVersion("OWC10", 1)]
		Int32 AddBitmap(Int32 iWidth, Int32 iHeight, stdole.IPicture pip, Int32 crMask);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pfln">NetOffice.OWC10Api.FieldListNode pfln</param>
		[SupportByVersion("OWC10", 1)]
		NetOffice.OWC10Api.FieldListNode GetNextSelected(NetOffice.OWC10Api.FieldListNode pfln);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void ClearSelection();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="iImage">Int32 iImage</param>
		/// <param name="iOverlay">Int32 iOverlay</param>
		[SupportByVersion("OWC10", 1)]
		void SetOverlayImage(Int32 iImage, Int32 iOverlay);

		#endregion
	}
}
