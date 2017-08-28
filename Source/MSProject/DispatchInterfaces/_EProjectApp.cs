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
 	public class _EProjectApp : _MSProject
	{
		#pragma warning disable

		#region Type Information

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

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _EProjectApp(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _EProjectApp(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _EProjectApp(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _EProjectApp(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _EProjectApp(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _EProjectApp(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _EProjectApp() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _EProjectApp(string progId) : base(progId)
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
		public void NewProject(NetOffice.MSProjectApi.Project pj)
		{
			 Factory.ExecuteMethod(this, "NewProject", pj);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tsk">NetOffice.MSProjectApi.Task tsk</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void ProjectBeforeTaskDelete(NetOffice.MSProjectApi.Task tsk, bool cancel)
		{
			 Factory.ExecuteMethod(this, "ProjectBeforeTaskDelete", tsk, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="res">NetOffice.MSProjectApi.Resource res</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void ProjectBeforeResourceDelete(NetOffice.MSProjectApi.Resource res, bool cancel)
		{
			 Factory.ExecuteMethod(this, "ProjectBeforeResourceDelete", res, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="asg">NetOffice.MSProjectApi.Assignment asg</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void ProjectBeforeAssignmentDelete(NetOffice.MSProjectApi.Assignment asg, bool cancel)
		{
			 Factory.ExecuteMethod(this, "ProjectBeforeAssignmentDelete", asg, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tsk">NetOffice.MSProjectApi.Task tsk</param>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="newVal">object newVal</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void ProjectBeforeTaskChange(NetOffice.MSProjectApi.Task tsk, NetOffice.MSProjectApi.Enums.PjField field, object newVal, bool cancel)
		{
			 Factory.ExecuteMethod(this, "ProjectBeforeTaskChange", tsk, field, newVal, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="res">NetOffice.MSProjectApi.Resource res</param>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField field</param>
		/// <param name="newVal">object newVal</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void ProjectBeforeResourceChange(NetOffice.MSProjectApi.Resource res, NetOffice.MSProjectApi.Enums.PjField field, object newVal, bool cancel)
		{
			 Factory.ExecuteMethod(this, "ProjectBeforeResourceChange", res, field, newVal, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="asg">NetOffice.MSProjectApi.Assignment asg</param>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjAssignmentField field</param>
		/// <param name="newVal">object newVal</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void ProjectBeforeAssignmentChange(NetOffice.MSProjectApi.Assignment asg, NetOffice.MSProjectApi.Enums.PjAssignmentField field, object newVal, bool cancel)
		{
			 Factory.ExecuteMethod(this, "ProjectBeforeAssignmentChange", asg, field, newVal, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void ProjectBeforeTaskNew(NetOffice.MSProjectApi.Project pj, bool cancel)
		{
			 Factory.ExecuteMethod(this, "ProjectBeforeTaskNew", pj, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void ProjectBeforeResourceNew(NetOffice.MSProjectApi.Project pj, bool cancel)
		{
			 Factory.ExecuteMethod(this, "ProjectBeforeResourceNew", pj, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void ProjectBeforeAssignmentNew(NetOffice.MSProjectApi.Project pj, bool cancel)
		{
			 Factory.ExecuteMethod(this, "ProjectBeforeAssignmentNew", pj, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void ProjectBeforeClose(NetOffice.MSProjectApi.Project pj, bool cancel)
		{
			 Factory.ExecuteMethod(this, "ProjectBeforeClose", pj, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void ProjectBeforePrint(NetOffice.MSProjectApi.Project pj, bool cancel)
		{
			 Factory.ExecuteMethod(this, "ProjectBeforePrint", pj, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="saveAsUi">bool saveAsUi</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void ProjectBeforeSave(NetOffice.MSProjectApi.Project pj, bool saveAsUi, bool cancel)
		{
			 Factory.ExecuteMethod(this, "ProjectBeforeSave", pj, saveAsUi, cancel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void ProjectCalculate(NetOffice.MSProjectApi.Project pj)
		{
			 Factory.ExecuteMethod(this, "ProjectCalculate", pj);
		}

		#endregion

		#pragma warning restore
	}
}
