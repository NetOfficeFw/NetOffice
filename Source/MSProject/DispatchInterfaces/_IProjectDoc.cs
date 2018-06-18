using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface _IProjectDoc 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00020B00-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.MSProjectApi.Project))]
    public interface _IProjectDoc : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Manager { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Company { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Author { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Keywords { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string ProjectNotes { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		object ProjectStart { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		object ProjectFinish { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		object CurrentDate { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		object StatusDate { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool ScheduleFromStart { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Comments { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Title { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Subject { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Windows Windows { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int16 MinuteLabelDisplay { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int16 HourLabelDisplay { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int16 DayLabelDisplay { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int16 WeekLabelDisplay { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int16 YearLabelDisplay { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int16 MonthLabelDisplay { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool SpaceBeforeTimeLabels { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjTaskFixedType DefaultTaskType { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool DefaultEffortDriven { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool UseFYStartYear { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool AutoFilter { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool HonorConstraints { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool MultipleCriticalPaths { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		object LevelFromDate { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		object LevelToDate { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool LevelEntireProject { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjAccrueAt DefaultFixedCostAccrual { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool SpreadCostsToStatusDate { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool SpreadPercentCompleteToStatusDate { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool AutoCalcCosts { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool ShowExternalSuccessors { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool ShowExternalPredecessors { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool ShowCrossProjectLinksInfo { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool AcceptNewExternalData { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjPhoneticType PhoneticType { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjWorkgroupMessages WorkgroupMessages { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string ServerURL { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string ServerPath { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool ReceiveNotifications { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool SendHyperlinkNote { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjColor HyperlinkColor { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjColor FollowedHyperlinkColor { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool UnderlineHyperlinks { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjTeamStatusCompletedWork AskForCompletedWork { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool TrackOvertimeWork { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool TeamMembersCanDeclineTasks { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool ShowEstimatedDuration { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool NewTasksEstimated { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool WBSCodeGenerate { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool WBSVerifyUniqueness { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool UpdateProjOnSave { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjAuthentication ServerIdentification { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool VBASigned { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool ExpandDatabaseTimephasedData { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		object DatabaseProjectUniqueID { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object ActualWork { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Cost1 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Cost2 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Cost3 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object BaselineWork { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object BaselineCost { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object FixedCost { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string WBS { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Delay { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Priority { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Duration1 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Duration2 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Duration3 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object PercentWorkComplete { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object FixedDuration { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object BaselineStart { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object BaselineFinish { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string Text1 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Start1 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Finish1 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string Text2 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Start2 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Finish2 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string Text3 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Start3 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Finish3 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string Text4 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Start4 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Finish4 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string Text5 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Start5 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Finish5 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string Text6 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string Text7 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string Text8 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string Text9 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string Text10 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Marked { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Flag1 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Flag2 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Flag3 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Flag4 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Flag5 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Flag6 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Flag7 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Flag8 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Flag9 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Flag10 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Rollup { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Double Number1 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Double Number2 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Double Number3 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Double Number4 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Double Number5 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string Notes { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string Contact { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object HideBar { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string CurrencySymbol { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjPlacement CurrencySymbolPosition { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int16 CurrencyDigits { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int32 ShowCriticalSlack { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjUnit DefaultDurationUnits { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjUnit DefaultWorkUnits { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool StartOnCurrentDate { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool AutoTrack { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool AutoSplitTasks { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool AutoLinkTasks { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		object DefaultStartTime { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		object DefaultFinishTime { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Double HoursPerDay { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Double HoursPerWeek { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Double DaysPerMonth { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		object DefaultResourceStandardRate { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		object DefaultResourceOvertimeRate { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool DisplayProjectSummaryTask { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool AutoAddResources { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjWeekday StartWeekOn { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjMonth StartYearIn { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool AllowTaskDelegation { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjPublishInformationOnSave PublishInformationOnSave { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string ProjectGuideFunctionalLayoutPage { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string ProjectGuideSaveBuffer { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string ProjectGuideContent { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool ProjectServerUsedForTracking { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjProjectServerTrackingMethod TrackingMethod { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool MoveCompleted { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool AndMoveRemaining { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool MoveRemaining { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool AndMoveCompleted { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjEarnedValueMethod DefaultEarnedValueMethod { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjBaselines EarnedValueBaseline { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool ProjectGuideUseDefaultFunctionalLayoutPage { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool ProjectGuideUseDefaultContent { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool EnterpriseActualsSynched { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool RemoveFileProperties { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool AdministrativeProject { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Windows2 Windows2 { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string _CodeName { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string CodeName { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Tasks OutlineChildren { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object CostVariance { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Task ProjectSummaryTask { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object RemainingCost { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object BCWP { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object BCWS { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object SV { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object CV { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string OutlineNumber { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Critical { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object FreeSlack { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object TotalSlack { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int32 UniqueID { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 OutlineLevel { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object BaselineDuration { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object DurationVariance { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object EarlyStart { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object EarlyFinish { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object LateStart { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object StartVariance { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object FinishVariance { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Project { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Milestone { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object RemainingDuration { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object PercentComplete { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Start { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Finish { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object ResourceNames { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object ResourceInitials { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Resume { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Stop { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object ResumeNoEarlierThan { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object ConstraintType { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object ConstraintDate { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object ActualCost { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Cost { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Created { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object ActualDuration { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Duration { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object LateFinish { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object ActualFinish { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 Objects { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object RemainingWork { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object ResourceGroup { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object ActualStart { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Summary { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string Template { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object UpdateNeeded { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Work { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object WorkVariance { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object LinkedFields { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object Confirmed { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool ReadOnly { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool HasPassword { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool WriteReserved { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		object Index { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.List MapList { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Tasks Tasks { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Resources Resources { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Calendars BaseCalendars { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14), ProxyResult]
		object BuiltinDocumentProperties { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14), ProxyResult]
		object CustomDocumentProperties { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14), ProxyResult]
		object Container { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Calendar Calendar { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int32 NumberOfTasks { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int32 NumberOfResources { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string FullName { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string Path { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string ResourcePoolName { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool Saved { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		object CreationDate { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		object LastPrintedDate { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		object LastSaveDate { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string LastSavedBy { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string RevisionNumber { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.List ViewList { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.List TaskViewList { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.List ResourceViewList { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool ReadOnlyRecommended { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.List ReportList { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.List TaskFilterList { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.List ResourceFilterList { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.List TaskTableList { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.List ResourceTableList { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string CurrentView { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string CurrentTable { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string CurrentFilter { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int32 ID { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.OfficeApi.CommandBars CommandBars { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool UserControl { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.VBIDEApi.VBProject VBProject { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Subprojects Subprojects { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string CurrentGroup { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.List TaskGroupList { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.List ResourceGroupList { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.TaskGroups TaskGroups { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.ResourceGroups ResourceGroups { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjProjectType Type { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[BaseResult]
		NetOffice.MSProjectApi.Views Views { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Tables TaskTables { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Tables ResourceTables { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Filters TaskFilters { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Filters ResourceFilters { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.ViewsSingle ViewsSingle { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.ViewsCombination ViewsCombination { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		/// <param name="baseline">NetOffice.MSProjectApi.Enums.PjBaselines baseline</param>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		object get_BaselineSavedDate(NetOffice.MSProjectApi.Enums.PjBaselines baseline);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Alias for get_BaselineSavedDate
		/// </summary>
		/// <param name="baseline">NetOffice.MSProjectApi.Enums.PjBaselines baseline</param>
		[SupportByVersion("MSProject", 11,12,14), Redirect("get_BaselineSavedDate")]
		object BaselineSavedDate(NetOffice.MSProjectApi.Enums.PjBaselines baseline);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string ProjectNamePrefix { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string VersionName { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int32 TempToDoList { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.OutlineCodes OutlineCodes { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.OfficeApi.SharedWorkspace SharedWorkspace { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.OfficeApi.DocumentLibraryVersions DocumentLibraryVersions { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool CanCheckIn { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string CurrencyCode { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int32 TaskErrorCount { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool IsTemplate { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		Int32 HyperlinkColorEx { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		Int32 FollowedHyperlinkColorEx { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		bool NewTasksCreatedAsManual { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		NetOffice.MSProjectApi.TaskGroups2 TaskGroups2 { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		NetOffice.MSProjectApi.ResourceGroups2 ResourceGroups2 { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		bool ManuallyScheduledTasksAutoRespectLinks { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		bool KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		bool ShowTaskWarnings { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		bool ShowTaskSuggestions { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		NetOffice.MSProjectApi.Tasks DetectCycle { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Reports Reports { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		bool IsCheckoutMsgBarVisible { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		bool IsCheckoutOSVisible { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		/// <param name="backup">optional object backup</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="taskInformation">optional object taskInformation</param>
		/// <param name="filtered">optional object filtered</param>
		/// <param name="table">optional object table</param>
		/// <param name="userID">optional object userID</param>
		/// <param name="databasePassWord">optional object databasePassWord</param>
		/// <param name="formatID">optional object formatID</param>
		/// <param name="map">optional object map</param>
		/// <param name="clearBaseline">optional object clearBaseline</param>
		/// <param name="clearActuals">optional object clearActuals</param>
		/// <param name="clearResourceRates">optional object clearResourceRates</param>
		/// <param name="clearFixedCosts">optional object clearFixedCosts</param>
		[SupportByVersion("MSProject", 11,12,14)]
		void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord, object formatID, object map, object clearBaseline, object clearActuals, object clearResourceRates, object clearFixedCosts);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		void SaveAs(object name);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		void SaveAs(object name, object format);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		/// <param name="backup">optional object backup</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		void SaveAs(object name, object format, object backup);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		/// <param name="backup">optional object backup</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		void SaveAs(object name, object format, object backup, object readOnly);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		/// <param name="backup">optional object backup</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="taskInformation">optional object taskInformation</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		void SaveAs(object name, object format, object backup, object readOnly, object taskInformation);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		/// <param name="backup">optional object backup</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="taskInformation">optional object taskInformation</param>
		/// <param name="filtered">optional object filtered</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		/// <param name="backup">optional object backup</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="taskInformation">optional object taskInformation</param>
		/// <param name="filtered">optional object filtered</param>
		/// <param name="table">optional object table</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		/// <param name="backup">optional object backup</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="taskInformation">optional object taskInformation</param>
		/// <param name="filtered">optional object filtered</param>
		/// <param name="table">optional object table</param>
		/// <param name="userID">optional object userID</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		/// <param name="backup">optional object backup</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="taskInformation">optional object taskInformation</param>
		/// <param name="filtered">optional object filtered</param>
		/// <param name="table">optional object table</param>
		/// <param name="userID">optional object userID</param>
		/// <param name="databasePassWord">optional object databasePassWord</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		/// <param name="backup">optional object backup</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="taskInformation">optional object taskInformation</param>
		/// <param name="filtered">optional object filtered</param>
		/// <param name="table">optional object table</param>
		/// <param name="userID">optional object userID</param>
		/// <param name="databasePassWord">optional object databasePassWord</param>
		/// <param name="formatID">optional object formatID</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord, object formatID);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		/// <param name="backup">optional object backup</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="taskInformation">optional object taskInformation</param>
		/// <param name="filtered">optional object filtered</param>
		/// <param name="table">optional object table</param>
		/// <param name="userID">optional object userID</param>
		/// <param name="databasePassWord">optional object databasePassWord</param>
		/// <param name="formatID">optional object formatID</param>
		/// <param name="map">optional object map</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord, object formatID, object map);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		/// <param name="backup">optional object backup</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="taskInformation">optional object taskInformation</param>
		/// <param name="filtered">optional object filtered</param>
		/// <param name="table">optional object table</param>
		/// <param name="userID">optional object userID</param>
		/// <param name="databasePassWord">optional object databasePassWord</param>
		/// <param name="formatID">optional object formatID</param>
		/// <param name="map">optional object map</param>
		/// <param name="clearBaseline">optional object clearBaseline</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord, object formatID, object map, object clearBaseline);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		/// <param name="backup">optional object backup</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="taskInformation">optional object taskInformation</param>
		/// <param name="filtered">optional object filtered</param>
		/// <param name="table">optional object table</param>
		/// <param name="userID">optional object userID</param>
		/// <param name="databasePassWord">optional object databasePassWord</param>
		/// <param name="formatID">optional object formatID</param>
		/// <param name="map">optional object map</param>
		/// <param name="clearBaseline">optional object clearBaseline</param>
		/// <param name="clearActuals">optional object clearActuals</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord, object formatID, object map, object clearBaseline, object clearActuals);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		/// <param name="backup">optional object backup</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="taskInformation">optional object taskInformation</param>
		/// <param name="filtered">optional object filtered</param>
		/// <param name="table">optional object table</param>
		/// <param name="userID">optional object userID</param>
		/// <param name="databasePassWord">optional object databasePassWord</param>
		/// <param name="formatID">optional object formatID</param>
		/// <param name="map">optional object map</param>
		/// <param name="clearBaseline">optional object clearBaseline</param>
		/// <param name="clearActuals">optional object clearActuals</param>
		/// <param name="clearResourceRates">optional object clearResourceRates</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord, object formatID, object map, object clearBaseline, object clearActuals, object clearResourceRates);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		void Activate();

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		void LevelClearDates();

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="value">string value</param>
		[SupportByVersion("MSProject", 11,12,14)]
		void AppendNotes(string value);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		void MakeServerURLTrusted();

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="comment">optional object comment</param>
		/// <param name="makePublic">optional object makePublic</param>
		[SupportByVersion("MSProject", 11,12,14)]
		void CheckIn(object saveChanges, object comment, object makePublic);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		void CheckIn();

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="saveChanges">optional object saveChanges</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		void CheckIn(object saveChanges);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="comment">optional object comment</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		void CheckIn(object saveChanges, object comment);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="objectType">NetOffice.MSProjectApi.Enums.PjOrganizer objectType</param>
		/// <param name="objectName">string objectName</param>
		[SupportByVersion("MSProject", 11,12,14)]
		string GetObjectMatchingID(NetOffice.MSProjectApi.Enums.PjOrganizer objectType, string objectName);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="objectType">NetOffice.MSProjectApi.Enums.PjOrganizer objectType</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="matchingID">string matchingID</param>
		[SupportByVersion("MSProject", 11,12,14)]
		void SetObjectMatchingID(NetOffice.MSProjectApi.Enums.PjOrganizer objectType, string objectName, string matchingID);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="objectType">NetOffice.MSProjectApi.Enums.PjOrganizer objectType</param>
		/// <param name="matchingID">string matchingID</param>
		[SupportByVersion("MSProject", 11,12,14)]
		string GetDisplayNameFromObjectMatchingID(NetOffice.MSProjectApi.Enums.PjOrganizer objectType, string matchingID);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableName">string deliverableName</param>
		/// <param name="deliverableStartDate">object deliverableStartDate</param>
		/// <param name="deliverableFinishDate">object deliverableFinishDate</param>
		/// <param name="taskGuid">string taskGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		string DeliverableCreate(string deliverableName, object deliverableStartDate, object deliverableFinishDate, string taskGuid);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		/// <param name="deliverableName">string deliverableName</param>
		/// <param name="deliverableStartDate">object deliverableStartDate</param>
		/// <param name="deliverableFinishDate">object deliverableFinishDate</param>
		[SupportByVersion("MSProject", 11,12,14)]
		bool DeliverableUpdate(string deliverableGuid, string deliverableName, object deliverableStartDate, object deliverableFinishDate);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		bool DeliverableDelete(string deliverableGuid);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		/// <param name="taskGuid">string taskGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		bool DeliverableDependencyCreate(string deliverableGuid, string taskGuid);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		bool DeliverableDependencyDelete(string deliverableGuid);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">optional object deliverableGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		bool DeliverableRefreshServerCache(object deliverableGuid);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		bool DeliverableRefreshServerCache();

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		object DeliverablesGetServerCachedXml();

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		object DeliverablesGetXml();

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string GetServerProjectGuid();

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		/// <param name="taskGuid">string taskGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		bool DeliverableLinkToTask(string deliverableGuid, string taskGuid);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		bool DeliverableLinkToProject(string deliverableGuid);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		bool DeliverablesClearAll();

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		bool DeliverableAcceptChanges(string deliverableGuid);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string DeliverablesGetProviderProjects();

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="projectGuid">string projectGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		object DeliverablesGetByProject(string projectGuid);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="taskGuid">string taskGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		Int32 GetTaskIndexByGuid(string taskGuid);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="projectGuid">string projectGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		object ReadWssData(string projectGuid);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		object GetWinprojURLs();

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int32 LocalResourceErrorCount();

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int32 ImportResourceErrorCount();

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int32 ResourceErrorCount();

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int32 LocalResourceCount();

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		Int32 ResourceCount();

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="assignmentToDrag">object assignmentToDrag</param>
		/// <param name="destinationResource">optional object destinationResource</param>
		/// <param name="destinationTime">optional object destinationTime</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSProject", 11,14)]
		void RSVDragSimulator(object assignmentToDrag, object destinationResource, object destinationTime);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="assignmentToDrag">object assignmentToDrag</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		void RSVDragSimulator(object assignmentToDrag);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="assignmentToDrag">object assignmentToDrag</param>
		/// <param name="destinationResource">optional object destinationResource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		void RSVDragSimulator(object assignmentToDrag, object destinationResource);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="customUIXML">string customUIXML</param>
		[SupportByVersion("MSProject", 11,14)]
		void SetCustomUI(string customUIXML);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileType">optional NetOffice.MSProjectApi.Enums.PjDocExportType FileType = 0</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="includeDocumentMarkup">optional bool IncludeDocumentMarkup = true</param>
		/// <param name="archiveFormat">optional bool ArchiveFormat = false</param>
		/// <param name="fromDate">optional object fromDate</param>
		/// <param name="toDate">optional object toDate</param>
		/// <param name="fixedFormatExtClassPtr">optional object fixedFormatExtClassPtr</param>
		[SupportByVersion("MSProject", 11,14)]
		void ExportAsFixedFormat(string filename, object fileType, object includeDocumentProperties, object includeDocumentMarkup, object archiveFormat, object fromDate, object toDate, object fixedFormatExtClassPtr);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		void ExportAsFixedFormat(string filename);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileType">optional NetOffice.MSProjectApi.Enums.PjDocExportType FileType = 0</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		void ExportAsFixedFormat(string filename, object fileType);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileType">optional NetOffice.MSProjectApi.Enums.PjDocExportType FileType = 0</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		void ExportAsFixedFormat(string filename, object fileType, object includeDocumentProperties);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileType">optional NetOffice.MSProjectApi.Enums.PjDocExportType FileType = 0</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="includeDocumentMarkup">optional bool IncludeDocumentMarkup = true</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		void ExportAsFixedFormat(string filename, object fileType, object includeDocumentProperties, object includeDocumentMarkup);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileType">optional NetOffice.MSProjectApi.Enums.PjDocExportType FileType = 0</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="includeDocumentMarkup">optional bool IncludeDocumentMarkup = true</param>
		/// <param name="archiveFormat">optional bool ArchiveFormat = false</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		void ExportAsFixedFormat(string filename, object fileType, object includeDocumentProperties, object includeDocumentMarkup, object archiveFormat);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileType">optional NetOffice.MSProjectApi.Enums.PjDocExportType FileType = 0</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="includeDocumentMarkup">optional bool IncludeDocumentMarkup = true</param>
		/// <param name="archiveFormat">optional bool ArchiveFormat = false</param>
		/// <param name="fromDate">optional object fromDate</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		void ExportAsFixedFormat(string filename, object fileType, object includeDocumentProperties, object includeDocumentMarkup, object archiveFormat, object fromDate);

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileType">optional NetOffice.MSProjectApi.Enums.PjDocExportType FileType = 0</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="includeDocumentMarkup">optional bool IncludeDocumentMarkup = true</param>
		/// <param name="archiveFormat">optional bool ArchiveFormat = false</param>
		/// <param name="fromDate">optional object fromDate</param>
		/// <param name="toDate">optional object toDate</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		void ExportAsFixedFormat(string filename, object fileType, object includeDocumentProperties, object includeDocumentMarkup, object archiveFormat, object fromDate, object toDate);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		Int32 CheckoutProject();

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		Int32 HideCheckoutMsgBar();

		#endregion
	}
}
