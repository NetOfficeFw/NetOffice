using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface PPListBox 
	/// SupportByVersion PowerPoint, 9
	/// </summary>
	[SupportByVersion("PowerPoint", 9)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("914934A8-5A91-11CF-8700-00AA0060263B")]
	public interface PPListBox : PPControl
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.PPStrings Strings { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.Enums.PpListBoxSelectionStyle SelectionStyle { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		Int32 FocusItem { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		Int32 TopItem { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		string OnSelectionChange { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		string OnDoubleClick { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get/Set
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("PowerPoint", 9)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.OfficeApi.Enums.MsoTriState get_IsSelected(Int32 index);

        /// <summary>
        /// SupportByVersion PowerPoint 9
        /// Get/Set
        /// </summary>
        /// <param name="index">Int32 index</param>
        /// <param name="value">NetOffice.OfficeApi.Enums.MsoTriState value</param>
        [SupportByVersion("PowerPoint", 9)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		void set_IsSelected(Int32 index, NetOffice.OfficeApi.Enums.MsoTriState value);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Alias for get_IsSelected
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("PowerPoint", 9), Redirect("get_IsSelected")]
		NetOffice.OfficeApi.Enums.MsoTriState IsSelected(Int32 index);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		NetOffice.PowerPointApi.Enums.PpListBoxAbbreviationStyle IsAbbreviated { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="safeArrayTabStops">object safeArrayTabStops</param>
		[SupportByVersion("PowerPoint", 9)]
		void SetTabStops(object safeArrayTabStops);

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="style">NetOffice.PowerPointApi.Enums.PpListBoxAbbreviationStyle style</param>
		[SupportByVersion("PowerPoint", 9)]
		void Abbreviate(NetOffice.PowerPointApi.Enums.PpListBoxAbbreviationStyle style);

		#endregion
	}
}
