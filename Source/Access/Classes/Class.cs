using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Class_InitializeEventHandler();
	public delegate void Class_TerminateEventHandler();
    #pragma warning restore

    #endregion

    /// <summary>
    /// CoClass Class
    /// SupportByVersion Access, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._DummyEvents))]
	[TypeId("8B06E321-B23C-11CF-89A8-00A0C9054129")]
    public interface Class : _Dummy, IEventBinding
    {
        #region Events

        /// <summary>
        /// SupportByVersion Access 9 10 11 12 14 15,16
        /// </summary>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        event Class_InitializeEventHandler InitializeEvent;

        /// <summary>
		/// SupportByVersion Access 9 10 11 12 14 15,16
		/// </summary>
		[SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        event Class_TerminateEventHandler TerminateEvent;

        #endregion
    }
}
