using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.MSProjectApi
{
	///<summary>
	/// DispatchInterface _EProjectApp 
	/// SupportByVersion MSProject, 11,12,14
	///</summary>
	[SupportByVersionAttribute("MSProject", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _EProjectApp : _MSProject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		/// 
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public void NewProject(NetOffice.MSProjectApi.Project pj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pj);
			Invoker.Method(this, "NewProject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="tsk">NetOffice.MSProjectApi.Task tsk</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public void ProjectBeforeTaskDelete(NetOffice.MSProjectApi.Task tsk, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tsk, cancel);
			Invoker.Method(this, "ProjectBeforeTaskDelete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="res">NetOffice.MSProjectApi.Resource res</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public void ProjectBeforeResourceDelete(NetOffice.MSProjectApi.Resource res, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(res, cancel);
			Invoker.Method(this, "ProjectBeforeResourceDelete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="asg">NetOffice.MSProjectApi.Assignment asg</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public void ProjectBeforeAssignmentDelete(NetOffice.MSProjectApi.Assignment asg, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(asg, cancel);
			Invoker.Method(this, "ProjectBeforeAssignmentDelete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="tsk">NetOffice.MSProjectApi.Task tsk</param>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField Field</param>
		/// <param name="newVal">object NewVal</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public void ProjectBeforeTaskChange(NetOffice.MSProjectApi.Task tsk, NetOffice.MSProjectApi.Enums.PjField field, object newVal, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tsk, field, newVal, cancel);
			Invoker.Method(this, "ProjectBeforeTaskChange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="res">NetOffice.MSProjectApi.Resource res</param>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjField Field</param>
		/// <param name="newVal">object NewVal</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public void ProjectBeforeResourceChange(NetOffice.MSProjectApi.Resource res, NetOffice.MSProjectApi.Enums.PjField field, object newVal, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(res, field, newVal, cancel);
			Invoker.Method(this, "ProjectBeforeResourceChange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="asg">NetOffice.MSProjectApi.Assignment asg</param>
		/// <param name="field">NetOffice.MSProjectApi.Enums.PjAssignmentField Field</param>
		/// <param name="newVal">object NewVal</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public void ProjectBeforeAssignmentChange(NetOffice.MSProjectApi.Assignment asg, NetOffice.MSProjectApi.Enums.PjAssignmentField field, object newVal, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(asg, field, newVal, cancel);
			Invoker.Method(this, "ProjectBeforeAssignmentChange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public void ProjectBeforeTaskNew(NetOffice.MSProjectApi.Project pj, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pj, cancel);
			Invoker.Method(this, "ProjectBeforeTaskNew", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public void ProjectBeforeResourceNew(NetOffice.MSProjectApi.Project pj, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pj, cancel);
			Invoker.Method(this, "ProjectBeforeResourceNew", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public void ProjectBeforeAssignmentNew(NetOffice.MSProjectApi.Project pj, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pj, cancel);
			Invoker.Method(this, "ProjectBeforeAssignmentNew", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public void ProjectBeforeClose(NetOffice.MSProjectApi.Project pj, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pj, cancel);
			Invoker.Method(this, "ProjectBeforeClose", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public void ProjectBeforePrint(NetOffice.MSProjectApi.Project pj, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pj, cancel);
			Invoker.Method(this, "ProjectBeforePrint", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		/// <param name="saveAsUi">bool SaveAsUi</param>
		/// <param name="cancel">bool Cancel</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public void ProjectBeforeSave(NetOffice.MSProjectApi.Project pj, bool saveAsUi, bool cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pj, saveAsUi, cancel);
			Invoker.Method(this, "ProjectBeforeSave", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="pj">NetOffice.MSProjectApi.Project pj</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public void ProjectCalculate(NetOffice.MSProjectApi.Project pj)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pj);
			Invoker.Method(this, "ProjectCalculate", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}