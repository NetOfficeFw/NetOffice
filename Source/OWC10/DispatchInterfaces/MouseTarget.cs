using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface MouseTarget 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("8F8E5640-CDB0-11D2-8F2E-00600893B533")]
	public interface MouseTarget : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="cursor">Int32 cursor</param>
		[SupportByVersion("OWC10", 1)]
		void MouseEnter(Int32 x, Int32 y, Int32 cursor);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="cursor">Int32 cursor</param>
		[SupportByVersion("OWC10", 1)]
		void MouseOver(Int32 x, Int32 y, Int32 cursor);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void MouseLeave();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="button">Int32 button</param>
		[SupportByVersion("OWC10", 1)]
		void MouseDown(Int32 x, Int32 y, Int32 button);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="button">Int32 button</param>
		[SupportByVersion("OWC10", 1)]
		void MouseUp(Int32 x, Int32 y, Int32 button);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="button">Int32 button</param>
		[SupportByVersion("OWC10", 1)]
		void MouseClick(Int32 x, Int32 y, Int32 button);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="button">Int32 button</param>
		[SupportByVersion("OWC10", 1)]
		void MouseDblClick(Int32 x, Int32 y, Int32 button);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="page">bool page</param>
		/// <param name="count">Int32 count</param>
		[SupportByVersion("OWC10", 1)]
		void MouseWheel(bool page, Int32 count);

		#endregion
	}
}
