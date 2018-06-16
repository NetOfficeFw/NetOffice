using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// Interface MouseDownHandler 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("914934BF-5A91-11CF-8700-00AA0060263B")]
	public interface MouseDownHandler : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="activeWin">object activeWin</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Int32 OnMouseDown(object activeWin);

		#endregion
	}
}
