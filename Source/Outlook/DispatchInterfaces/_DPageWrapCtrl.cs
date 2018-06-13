using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _DPageWrapCtrl 
	/// SupportByVersion Outlook, 10
	/// </summary>
	[SupportByVersion("Outlook", 10)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("0006F096-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi._PageWrapCtrl))]
    public interface _DPageWrapCtrl : ICOMObject
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
