using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSProjectApi;

namespace NetOffice.MSProjectApi.Behind
{
	/// <summary>
	/// DispatchInterface _EProjectApp 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class _EProjectApp : _MSProject, NetOffice.MSProjectApi._EProjectApp
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.MSProjectApi._EProjectApp);
                return _contractType;
            }
        }
        private static Type _contractType;


		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(_EProjectApp);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _EProjectApp() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void NewProject(NetOffice.MSProjectApi.Project pj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NewProject", pj);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tsk">NetOffice.MSProjectApi.Task tsk</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void ProjectBeforeTaskDelete(NetOffice.MSProjectApi.Task tsk, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ProjectBeforeTaskDelete", tsk, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="res">NetOffice.MSProjectApi.Resource res</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void ProjectBeforeResourceDelete(NetOffice.MSProjectApi.Resource res, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ProjectBeforeResourceDelete", res, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="asg">NetOffice.MSProjectApi.Assignment asg</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void ProjectBeforeAssignmentDelete(NetOffice.MSProjectApi.Assignment asg, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ProjectBeforeAssignmentDelete", asg, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tsk">NetOffice.MSProjectApi.Task tsk</param>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="newVal">object newVal</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void ProjectBeforeTaskChange(NetOffice.MSProjectApi.Task tsk, NetOffice.MSProjectApi.Enums.PjField field, object newVal, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ProjectBeforeTaskChange", tsk, field, newVal, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="res">NetOffice.MSProjectApi.Resource res</param>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="newVal">object newVal</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void ProjectBeforeResourceChange(NetOffice.MSProjectApi.Resource res, NetOffice.MSProjectApi.Enums.PjField field, object newVal, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ProjectBeforeResourceChange", res, field, newVal, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="asg">NetOffice.MSProjectApi.Assignment asg</param>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjAssignmentField field</param>
		/// <param name="newVal">object newVal</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void ProjectBeforeAssignmentChange(NetOffice.MSProjectApi.Assignment asg, NetOffice.MSProjectApi.Enums.PjAssignmentField field, object newVal, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ProjectBeforeAssignmentChange", asg, field, newVal, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void ProjectBeforeTaskNew(NetOffice.MSProjectApi.Project pj, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ProjectBeforeTaskNew", pj, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void ProjectBeforeResourceNew(NetOffice.MSProjectApi.Project pj, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ProjectBeforeResourceNew", pj, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void ProjectBeforeAssignmentNew(NetOffice.MSProjectApi.Project pj, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ProjectBeforeAssignmentNew", pj, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void ProjectBeforeClose(NetOffice.MSProjectApi.Project pj, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ProjectBeforeClose", pj, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void ProjectBeforePrint(NetOffice.MSProjectApi.Project pj, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ProjectBeforePrint", pj, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="saveAsUi">bool saveAsUi</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void ProjectBeforeSave(NetOffice.MSProjectApi.Project pj, bool saveAsUi, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ProjectBeforeSave", pj, saveAsUi, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void ProjectCalculate(NetOffice.MSProjectApi.Project pj)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ProjectCalculate", pj);
		}

		#endregion

		#pragma warning restore
	}
}

