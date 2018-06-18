using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
	#region Delegates

	#pragma warning disable
	public delegate void Project_OpenEventHandler(NetOffice.MSProjectApi.Project pj);
	public delegate void Project_BeforeCloseEventHandler(NetOffice.MSProjectApi.Project pj);
	public delegate void Project_BeforeSaveEventHandler(NetOffice.MSProjectApi.Project pj);
	public delegate void Project_BeforePrintEventHandler(NetOffice.MSProjectApi.Project pj);
	public delegate void Project_CalculateEventHandler(NetOffice.MSProjectApi.Project pj);
	public delegate void Project_ChangeEventHandler(NetOffice.MSProjectApi.Project pj);
	public delegate void Project_ActivateEventHandler(NetOffice.MSProjectApi.Project pj);
	public delegate void Project_DeactivateEventHandler(NetOffice.MSProjectApi.Project pj);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass Project 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920664(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(EventContracts._EProjectDoc))]
	[TypeId("1019A320-508A-11CF-A49D-00AA00574C74")]
    public interface Project : _IProjectDoc, IEventBinding
	{
		#region Events

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Project_OpenEventHandler OpenEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Project_BeforeCloseEventHandler BeforeCloseEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Project_BeforeSaveEventHandler BeforeSaveEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Project_BeforePrintEventHandler BeforePrintEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Project_CalculateEventHandler CalculateEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Project_ChangeEventHandler ChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Project_ActivateEventHandler ActivateEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		event Project_DeactivateEventHandler DeactivateEvent;

		#endregion
	}
}
