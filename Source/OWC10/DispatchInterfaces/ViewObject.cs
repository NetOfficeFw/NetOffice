using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface ViewObject 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("8C950530-C84C-11D2-8F2D-00600893B533")]
	public interface ViewObject : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object get_Query(Int32 x, Int32 y);

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_Query
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="x">Int32 x</param>
		/// <param name="y">Int32 y</param>
		[SupportByVersion("OWC10", 1), ProxyResult, Redirect("get_Query")]
		object Query(Int32 x, Int32 y);

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="surface">NetOffice.OWC10Api.ViewSurface surface</param>
		/// <param name="cx1">Int32 cx1</param>
		/// <param name="cy1">Int32 cy1</param>
		/// <param name="cx2">Int32 cx2</param>
		/// <param name="cy2">Int32 cy2</param>
		/// <param name="left">Int32 left</param>
		/// <param name="top">Int32 top</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		[SupportByVersion("OWC10", 1)]
		void Draw(NetOffice.OWC10Api.ViewSurface surface, Int32 cx1, Int32 cy1, Int32 cx2, Int32 cy2, Int32 left, Int32 top, Int32 width, Int32 height);

		#endregion
	}
}
