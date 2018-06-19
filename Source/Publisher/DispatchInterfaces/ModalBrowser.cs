using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface ModalBrowser 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("FAC601B6-4A1A-4F69-9ABD-4B4DA640B2DB")]
	public interface ModalBrowser : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void TaskCompleted();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="lWidth">Int32 lWidth</param>
		/// <param name="lHeight">Int32 lHeight</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void ResizeTo(Int32 lWidth, Int32 lHeight);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="lx">Int32 lx</param>
		/// <param name="ly">Int32 ly</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void MoveTo(Int32 lx, Int32 ly);

		#endregion
	}
}
