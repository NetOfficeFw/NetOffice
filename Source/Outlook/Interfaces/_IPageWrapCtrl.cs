using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// Interface _IPageWrapCtrl 
	/// SupportByVersion Outlook, 10
	/// </summary>
	[SupportByVersion("Outlook", 10)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("494F0970-DD96-11D2-AF70-006008AFF117")]
	public interface _IPageWrapCtrl : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 10
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 10)]
		Int32 BackColor { get; set; }

		#endregion

	}
}
