using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// Interface MouseTracker 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("914934BE-5A91-11CF-8700-00AA0060263B")]
	public interface MouseTracker : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="x">Single x</param>
		/// <param name="y">Single y</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Int32 OnTrack(Single x, Single y);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="x">Single x</param>
		/// <param name="y">Single y</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Int32 EndTrack(Single x, Single y);

		#endregion
	}
}
