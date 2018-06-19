using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface ColorScheme 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("821941D8-F6DD-11D3-907C-00C04F799E3F")]
	public interface ColorScheme : ICOMObject
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
		/// </summary>
		/// <param name="colorIndex">NetOffice.PublisherApi.Enums.PbSchemeColorIndex colorIndex</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.PublisherApi.ColorFormat get_Colors(NetOffice.PublisherApi.Enums.PbSchemeColorIndex colorIndex);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Alias for get_Colors
		/// </summary>
		/// <param name="colorIndex">NetOffice.PublisherApi.Enums.PbSchemeColorIndex colorIndex</param>
		[SupportByVersion("Publisher", 14,15,16), Redirect("get_Colors")]
		NetOffice.PublisherApi.ColorFormat Colors(NetOffice.PublisherApi.Enums.PbSchemeColorIndex colorIndex);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		object Parent { get; }

		#endregion

	}
}
