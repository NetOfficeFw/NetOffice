using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface _EProjectApp 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("7B7597D0-0C9F-11D0-8C43-00A0C904DCD4")]
	public interface _EProjectApp : _MSProject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		[SupportByVersion("MSProject", 11,12,14)]
		void NewProject(NetOffice.MSProjectApi.Project pj);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tsk">NetOffice.MSProjectApi.Task tsk</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		void ProjectBeforeTaskDelete(NetOffice.MSProjectApi.Task tsk, bool cancel);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="res">NetOffice.MSProjectApi.Resource res</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		void ProjectBeforeResourceDelete(NetOffice.MSProjectApi.Resource res, bool cancel);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="asg">NetOffice.MSProjectApi.Assignment asg</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		void ProjectBeforeAssignmentDelete(NetOffice.MSProjectApi.Assignment asg, bool cancel);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tsk">NetOffice.MSProjectApi.Task tsk</param>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="newVal">object newVal</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		void ProjectBeforeTaskChange(NetOffice.MSProjectApi.Task tsk, NetOffice.MSProjectApi.Enums.PjField field, object newVal, bool cancel);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="res">NetOffice.MSProjectApi.Resource res</param>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="newVal">object newVal</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		void ProjectBeforeResourceChange(NetOffice.MSProjectApi.Resource res, NetOffice.MSProjectApi.Enums.PjField field, object newVal, bool cancel);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="asg">NetOffice.MSProjectApi.Assignment asg</param>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjAssignmentField field</param>
		/// <param name="newVal">object newVal</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		void ProjectBeforeAssignmentChange(NetOffice.MSProjectApi.Assignment asg, NetOffice.MSProjectApi.Enums.PjAssignmentField field, object newVal, bool cancel);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		void ProjectBeforeTaskNew(NetOffice.MSProjectApi.Project pj, bool cancel);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		void ProjectBeforeResourceNew(NetOffice.MSProjectApi.Project pj, bool cancel);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		void ProjectBeforeAssignmentNew(NetOffice.MSProjectApi.Project pj, bool cancel);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		void ProjectBeforeClose(NetOffice.MSProjectApi.Project pj, bool cancel);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		void ProjectBeforePrint(NetOffice.MSProjectApi.Project pj, bool cancel);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="saveAsUi">bool saveAsUi</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		void ProjectBeforeSave(NetOffice.MSProjectApi.Project pj, bool saveAsUi, bool cancel);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		[SupportByVersion("MSProject", 11,12,14)]
		void ProjectCalculate(NetOffice.MSProjectApi.Project pj);

		#endregion
	}
}
