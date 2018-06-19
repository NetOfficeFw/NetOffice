using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface ICagNotifySink 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("00021293-0000-0000-C000-000000000046")]
	public interface ICagNotifySink : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pClipMoniker">object pClipMoniker</param>
		/// <param name="pItemMoniker">object pItemMoniker</param>
		[SupportByVersion("Publisher", 14,15,16)]
		object InsertClip(object pClipMoniker, object pItemMoniker);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		object WindowIsClosing();

		#endregion
	}
}
