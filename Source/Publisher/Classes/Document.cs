using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Document_OpenEventHandler();
	public delegate void Document_BeforeCloseEventHandler(ref bool cancel);
	public delegate void Document_ShapesAddedEventHandler();
	public delegate void Document_WizardAfterChangeEventHandler();
	public delegate void Document_ShapesRemovedEventHandler();
	public delegate void Document_UndoEventHandler();
	public delegate void Document_RedoEventHandler();
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Document 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts.DocumentEvents))]
	[TypeId("09FD2EFF-5669-11D3-B65F-00C04F8EF32D")]
    public interface Document : _Document, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Document_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Document_BeforeCloseEventHandler BeforeCloseEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Document_ShapesAddedEventHandler ShapesAddedEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Document_WizardAfterChangeEventHandler WizardAfterChangeEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Document_ShapesRemovedEventHandler ShapesRemovedEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Document_UndoEventHandler UndoEvent;

		/// <summary>
		/// SupportByVersion Publisher 14 15 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		event Document_RedoEventHandler RedoEvent;

		#endregion
	}
}
