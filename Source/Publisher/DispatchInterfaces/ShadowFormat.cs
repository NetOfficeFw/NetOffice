using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface ShadowFormat 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("0002126C-0000-0000-C000-000000000046")]
	public interface ShadowFormat : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ColorFormat ForeColor { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState Obscured { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		object OffsetX { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		object OffsetY { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Single Transparency { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoShadowType Type { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState Visible { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		Single Blur { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		Single Size { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState RotateWithShape { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="increment">object increment</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void IncrementOffsetX(object increment);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="increment">object increment</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void IncrementOffsetY(object increment);

		#endregion
	}
}
