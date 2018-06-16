using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface DropTarget 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("77186960-CDB1-11D2-8F2E-00600893B533")]
	public interface DropTarget : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="keyState">Int32 keyState</param>
		/// <param name="effect">Int32 effect</param>
		/// <param name="_object">object object</param>
		[SupportByVersion("OWC10", 1)]
		void DragEnter(Int32 x, Int32 y, Int32 keyState, Int32 effect, object _object);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="keyState">Int32 keyState</param>
		/// <param name="effect">Int32 effect</param>
		[SupportByVersion("OWC10", 1)]
		void DragOver(Int32 x, Int32 y, Int32 keyState, Int32 effect);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		void DragLeave();

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		/// <param name="keyState">Int32 keyState</param>
		/// <param name="effect">Int32 effect</param>
		[SupportByVersion("OWC10", 1)]
		void Drop(Int32 x, Int32 y, Int32 keyState, Int32 effect);

		#endregion
	}
}
