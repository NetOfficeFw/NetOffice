using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void SubForm_EnterEventHandler();
	public delegate void SubForm_ExitEventHandler(ref Int16 cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass SubForm 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194842.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._SubFormEvents), typeof(EventContracts.DispSubFormEvents))]
	[TypeId("3B06E963-E47C-11CD-8701-00AA003F0F07")]
    public interface SubForm : _SubForm, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198039.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event SubForm_EnterEventHandler EnterEvent;

		/// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836988.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		event SubForm_ExitEventHandler ExitEvent;

		#endregion
	}
}
