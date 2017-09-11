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
 	public class _IProjectDoc : COMObject
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
                    _type = typeof(_IProjectDoc);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _IProjectDoc(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _IProjectDoc(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _IProjectDoc(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _IProjectDoc(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _IProjectDoc(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _IProjectDoc(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _IProjectDoc() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _IProjectDoc(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Manager
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Manager");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Manager", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Company
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Company");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Company", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Author
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Author");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Author", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Keywords
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Keywords");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Keywords", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string ProjectNotes
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ProjectNotes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ProjectNotes", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object ProjectStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ProjectStart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ProjectStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object ProjectFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ProjectFinish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ProjectFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object CurrentDate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CurrentDate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "CurrentDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object StatusDate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "StatusDate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "StatusDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool ScheduleFromStart
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ScheduleFromStart");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScheduleFromStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Comments
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Comments");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Comments", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Title
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Title");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Title", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Subject
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Subject");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Subject", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Windows Windows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Windows>(this, "Windows", NetOffice.MSProjectApi.Windows.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Windows", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int16 MinuteLabelDisplay
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "MinuteLabelDisplay");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MinuteLabelDisplay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int16 HourLabelDisplay
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "HourLabelDisplay");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HourLabelDisplay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int16 DayLabelDisplay
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "DayLabelDisplay");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DayLabelDisplay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int16 WeekLabelDisplay
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "WeekLabelDisplay");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WeekLabelDisplay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int16 YearLabelDisplay
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "YearLabelDisplay");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "YearLabelDisplay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int16 MonthLabelDisplay
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "MonthLabelDisplay");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MonthLabelDisplay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool SpaceBeforeTimeLabels
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SpaceBeforeTimeLabels");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SpaceBeforeTimeLabels", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjTaskFixedType DefaultTaskType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjTaskFixedType>(this, "DefaultTaskType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultTaskType", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool DefaultEffortDriven
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DefaultEffortDriven");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultEffortDriven", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool UseFYStartYear
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseFYStartYear");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseFYStartYear", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool AutoFilter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFilter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool HonorConstraints
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HonorConstraints");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HonorConstraints", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool MultipleCriticalPaths
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MultipleCriticalPaths");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MultipleCriticalPaths", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object LevelFromDate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LevelFromDate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "LevelFromDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object LevelToDate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LevelToDate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "LevelToDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool LevelEntireProject
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LevelEntireProject");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LevelEntireProject", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjAccrueAt DefaultFixedCostAccrual
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "DefaultFixedCostAccrual");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultFixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool SpreadCostsToStatusDate
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SpreadCostsToStatusDate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SpreadCostsToStatusDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool SpreadPercentCompleteToStatusDate
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SpreadPercentCompleteToStatusDate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SpreadPercentCompleteToStatusDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool AutoCalcCosts
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoCalcCosts");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoCalcCosts", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool ShowExternalSuccessors
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowExternalSuccessors");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowExternalSuccessors", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool ShowExternalPredecessors
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowExternalPredecessors");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowExternalPredecessors", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool ShowCrossProjectLinksInfo
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowCrossProjectLinksInfo");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowCrossProjectLinksInfo", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool AcceptNewExternalData
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AcceptNewExternalData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AcceptNewExternalData", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjPhoneticType PhoneticType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjPhoneticType>(this, "PhoneticType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PhoneticType", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjWorkgroupMessages WorkgroupMessages
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjWorkgroupMessages>(this, "WorkgroupMessages");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "WorkgroupMessages", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string ServerURL
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ServerURL");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ServerURL", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string ServerPath
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ServerPath");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ServerPath", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool ReceiveNotifications
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReceiveNotifications");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReceiveNotifications", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool SendHyperlinkNote
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SendHyperlinkNote");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SendHyperlinkNote", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjColor HyperlinkColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjColor>(this, "HyperlinkColor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "HyperlinkColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjColor FollowedHyperlinkColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjColor>(this, "FollowedHyperlinkColor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FollowedHyperlinkColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool UnderlineHyperlinks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UnderlineHyperlinks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UnderlineHyperlinks", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjTeamStatusCompletedWork AskForCompletedWork
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjTeamStatusCompletedWork>(this, "AskForCompletedWork");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AskForCompletedWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool TrackOvertimeWork
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TrackOvertimeWork");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TrackOvertimeWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool TeamMembersCanDeclineTasks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TeamMembersCanDeclineTasks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TeamMembersCanDeclineTasks", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool ShowEstimatedDuration
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowEstimatedDuration");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowEstimatedDuration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool NewTasksEstimated
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "NewTasksEstimated");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NewTasksEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool WBSCodeGenerate
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WBSCodeGenerate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WBSCodeGenerate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool WBSVerifyUniqueness
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WBSVerifyUniqueness");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WBSVerifyUniqueness", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool UpdateProjOnSave
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UpdateProjOnSave");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UpdateProjOnSave", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjAuthentication ServerIdentification
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAuthentication>(this, "ServerIdentification");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ServerIdentification", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool VBASigned
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "VBASigned");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "VBASigned", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool ExpandDatabaseTimephasedData
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ExpandDatabaseTimephasedData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ExpandDatabaseTimephasedData", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object DatabaseProjectUniqueID
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DatabaseProjectUniqueID");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DatabaseProjectUniqueID", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ActualWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ActualWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ActualWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Cost1
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Cost1");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Cost1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Cost2
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Cost2");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Cost2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Cost3
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Cost3");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Cost3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object BaselineWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BaselineWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BaselineWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object BaselineCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BaselineCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BaselineCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object FixedCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FixedCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string WBS
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "WBS");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WBS", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Delay
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Delay");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Delay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Priority
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Priority");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Priority", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Duration1
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration1");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Duration2
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration2");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Duration3
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration3");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PercentWorkComplete
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PercentWorkComplete");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PercentWorkComplete", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object FixedDuration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FixedDuration");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "FixedDuration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object BaselineStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BaselineStart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BaselineStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object BaselineFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BaselineFinish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BaselineFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Text1
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text1");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Start1
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Start1");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Start1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Finish1
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Finish1");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Finish1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Text2
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text2");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Start2
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Start2");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Start2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Finish2
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Finish2");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Finish2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Text3
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text3");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Start3
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Start3");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Start3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Finish3
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Finish3");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Finish3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Text4
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text4");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Start4
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Start4");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Start4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Finish4
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Finish4");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Finish4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Text5
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text5");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Start5
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Start5");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Start5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Finish5
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Finish5");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Finish5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Text6
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text6");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Text7
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text7");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Text8
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text8");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Text9
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text9");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Text10
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text10");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Marked
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Marked");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Marked", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Flag1
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag1");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Flag2
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag2");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Flag3
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag3");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Flag4
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag4");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Flag5
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag5");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Flag6
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag6");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Flag7
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag7");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Flag8
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag8");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Flag9
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag9");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Flag10
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag10");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Rollup
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Rollup");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Rollup", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double Number1
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number1");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double Number2
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number2");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double Number3
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number3");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double Number4
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number4");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double Number5
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number5");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Notes
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Notes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Notes", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Contact
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Contact");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Contact", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object HideBar
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "HideBar");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "HideBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string CurrencySymbol
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CurrencySymbol");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CurrencySymbol", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjPlacement CurrencySymbolPosition
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjPlacement>(this, "CurrencySymbolPosition");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CurrencySymbolPosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int16 CurrencyDigits
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "CurrencyDigits");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CurrencyDigits", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 ShowCriticalSlack
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ShowCriticalSlack");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowCriticalSlack", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjUnit DefaultDurationUnits
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjUnit>(this, "DefaultDurationUnits");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultDurationUnits", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjUnit DefaultWorkUnits
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjUnit>(this, "DefaultWorkUnits");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultWorkUnits", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool StartOnCurrentDate
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "StartOnCurrentDate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StartOnCurrentDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool AutoTrack
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoTrack");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoTrack", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool AutoSplitTasks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoSplitTasks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoSplitTasks", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool AutoLinkTasks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoLinkTasks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoLinkTasks", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object DefaultStartTime
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultStartTime");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DefaultStartTime", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object DefaultFinishTime
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultFinishTime");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DefaultFinishTime", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double HoursPerDay
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "HoursPerDay");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HoursPerDay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double HoursPerWeek
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "HoursPerWeek");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HoursPerWeek", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double DaysPerMonth
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "DaysPerMonth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DaysPerMonth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object DefaultResourceStandardRate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultResourceStandardRate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DefaultResourceStandardRate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object DefaultResourceOvertimeRate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultResourceOvertimeRate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DefaultResourceOvertimeRate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool DisplayProjectSummaryTask
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayProjectSummaryTask");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayProjectSummaryTask", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool AutoAddResources
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoAddResources");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoAddResources", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjWeekday StartWeekOn
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjWeekday>(this, "StartWeekOn");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "StartWeekOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjMonth StartYearIn
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjMonth>(this, "StartYearIn");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "StartYearIn", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool AllowTaskDelegation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowTaskDelegation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowTaskDelegation", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjPublishInformationOnSave PublishInformationOnSave
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjPublishInformationOnSave>(this, "PublishInformationOnSave");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PublishInformationOnSave", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string ProjectGuideFunctionalLayoutPage
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ProjectGuideFunctionalLayoutPage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ProjectGuideFunctionalLayoutPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string ProjectGuideSaveBuffer
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ProjectGuideSaveBuffer");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ProjectGuideSaveBuffer", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string ProjectGuideContent
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ProjectGuideContent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ProjectGuideContent", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool ProjectServerUsedForTracking
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ProjectServerUsedForTracking");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ProjectServerUsedForTracking", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjProjectServerTrackingMethod TrackingMethod
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjProjectServerTrackingMethod>(this, "TrackingMethod");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TrackingMethod", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool MoveCompleted
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MoveCompleted");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MoveCompleted", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool AndMoveRemaining
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AndMoveRemaining");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AndMoveRemaining", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool MoveRemaining
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MoveRemaining");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MoveRemaining", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool AndMoveCompleted
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AndMoveCompleted");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AndMoveCompleted", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjEarnedValueMethod DefaultEarnedValueMethod
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjEarnedValueMethod>(this, "DefaultEarnedValueMethod");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultEarnedValueMethod", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjBaselines EarnedValueBaseline
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjBaselines>(this, "EarnedValueBaseline");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "EarnedValueBaseline", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool ProjectGuideUseDefaultFunctionalLayoutPage
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ProjectGuideUseDefaultFunctionalLayoutPage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ProjectGuideUseDefaultFunctionalLayoutPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool ProjectGuideUseDefaultContent
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ProjectGuideUseDefaultContent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ProjectGuideUseDefaultContent", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool EnterpriseActualsSynched
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnterpriseActualsSynched");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseActualsSynched", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool RemoveFileProperties
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RemoveFileProperties");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RemoveFileProperties", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool AdministrativeProject
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AdministrativeProject");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AdministrativeProject", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Windows2 Windows2
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Windows2>(this, "Windows2", NetOffice.MSProjectApi.Windows2.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Windows2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string _CodeName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "_CodeName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "_CodeName", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string CodeName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CodeName");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Tasks OutlineChildren
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Tasks>(this, "OutlineChildren", NetOffice.MSProjectApi.Tasks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object CostVariance
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CostVariance");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Task ProjectSummaryTask
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Task>(this, "ProjectSummaryTask", NetOffice.MSProjectApi.Task.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object RemainingCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "RemainingCost");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object BCWP
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BCWP");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object BCWS
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BCWS");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object SV
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "SV");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object CV
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CV");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OutlineNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OutlineNumber");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Critical
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Critical");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object FreeSlack
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FreeSlack");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object TotalSlack
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TotalSlack");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 UniqueID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "UniqueID");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 OutlineLevel
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "OutlineLevel");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object BaselineDuration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BaselineDuration");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object DurationVariance
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DurationVariance");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object EarlyStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EarlyStart");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object EarlyFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EarlyFinish");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object LateStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LateStart");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object StartVariance
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "StartVariance");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object FinishVariance
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FinishVariance");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Project
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Project");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Milestone
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Milestone");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object RemainingDuration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "RemainingDuration");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PercentComplete
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PercentComplete");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Start
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Start");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Finish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Finish");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ResourceNames
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ResourceNames");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ResourceInitials
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ResourceInitials");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Resume
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Resume");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Stop
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Stop");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ResumeNoEarlierThan
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ResumeNoEarlierThan");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ConstraintType
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ConstraintType");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ConstraintDate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ConstraintDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ActualCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ActualCost");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Cost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Cost");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Created
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Created");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ActualDuration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ActualDuration");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Duration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object LateFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LateFinish");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ActualFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ActualFinish");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 Objects
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Objects");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object RemainingWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "RemainingWork");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ResourceGroup
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ResourceGroup");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ActualStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ActualStart");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Summary
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Summary");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Template
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Template");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object UpdateNeeded
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "UpdateNeeded");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Work
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Work");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object WorkVariance
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "WorkVariance");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object LinkedFields
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LinkedFields");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Confirmed
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Confirmed");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool ReadOnly
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReadOnly");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool HasPassword
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasPassword");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool WriteReserved
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WriteReserved");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Application>(this, "Application", NetOffice.MSProjectApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Index
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.List MapList
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "MapList", NetOffice.MSProjectApi.List.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Tasks Tasks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Tasks>(this, "Tasks", NetOffice.MSProjectApi.Tasks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Resources Resources
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Resources>(this, "Resources", NetOffice.MSProjectApi.Resources.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Calendars BaseCalendars
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Calendars>(this, "BaseCalendars", NetOffice.MSProjectApi.Calendars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14), ProxyResult]
		public object BuiltinDocumentProperties
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "BuiltinDocumentProperties");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14), ProxyResult]
		public object CustomDocumentProperties
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "CustomDocumentProperties");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14), ProxyResult]
		public object Container
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Container");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Calendar Calendar
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Calendar>(this, "Calendar", NetOffice.MSProjectApi.Calendar.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 NumberOfTasks
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "NumberOfTasks");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 NumberOfResources
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "NumberOfResources");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string FullName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FullName");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Path
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Path");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string ResourcePoolName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ResourcePoolName");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool Saved
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Saved");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object CreationDate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CreationDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object LastPrintedDate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LastPrintedDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object LastSaveDate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LastSaveDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string LastSavedBy
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LastSavedBy");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string RevisionNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RevisionNumber");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.List ViewList
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "ViewList", NetOffice.MSProjectApi.List.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.List TaskViewList
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "TaskViewList", NetOffice.MSProjectApi.List.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.List ResourceViewList
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "ResourceViewList", NetOffice.MSProjectApi.List.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool ReadOnlyRecommended
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReadOnlyRecommended");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.List ReportList
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "ReportList", NetOffice.MSProjectApi.List.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.List TaskFilterList
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "TaskFilterList", NetOffice.MSProjectApi.List.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.List ResourceFilterList
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "ResourceFilterList", NetOffice.MSProjectApi.List.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.List TaskTableList
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "TaskTableList", NetOffice.MSProjectApi.List.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.List ResourceTableList
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "ResourceTableList", NetOffice.MSProjectApi.List.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string CurrentView
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CurrentView");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string CurrentTable
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CurrentTable");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string CurrentFilter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CurrentFilter");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 ID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ID");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(this, "CommandBars", NetOffice.OfficeApi.CommandBars.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool UserControl
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UserControl");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.VBIDEApi.VBProject VBProject
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBProject>(this, "VBProject", NetOffice.VBIDEApi.VBProject.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Subprojects Subprojects
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Subprojects>(this, "Subprojects", NetOffice.MSProjectApi.Subprojects.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string CurrentGroup
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CurrentGroup");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.List TaskGroupList
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "TaskGroupList", NetOffice.MSProjectApi.List.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.List ResourceGroupList
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "ResourceGroupList", NetOffice.MSProjectApi.List.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TaskGroups TaskGroups
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.TaskGroups>(this, "TaskGroups", NetOffice.MSProjectApi.TaskGroups.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ResourceGroups ResourceGroups
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.ResourceGroups>(this, "ResourceGroups", NetOffice.MSProjectApi.ResourceGroups.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjProjectType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjProjectType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[BaseResult]
		public NetOffice.MSProjectApi.Views Views
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSProjectApi.Views>(this, "Views");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Tables TaskTables
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Tables>(this, "TaskTables", NetOffice.MSProjectApi.Tables.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Tables ResourceTables
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Tables>(this, "ResourceTables", NetOffice.MSProjectApi.Tables.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Filters TaskFilters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Filters>(this, "TaskFilters", NetOffice.MSProjectApi.Filters.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Filters ResourceFilters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Filters>(this, "ResourceFilters", NetOffice.MSProjectApi.Filters.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ViewsSingle ViewsSingle
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.ViewsSingle>(this, "ViewsSingle", NetOffice.MSProjectApi.ViewsSingle.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ViewsCombination ViewsCombination
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.ViewsCombination>(this, "ViewsCombination", NetOffice.MSProjectApi.ViewsCombination.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		/// <param name="baseline">NetOffice.MSProjectApi.Enums.PjBaselines baseline</param>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_BaselineSavedDate(NetOffice.MSProjectApi.Enums.PjBaselines baseline)
		{
			return Factory.ExecuteVariantPropertyGet(this, "BaselineSavedDate", baseline);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Alias for get_BaselineSavedDate
		/// </summary>
		/// <param name="baseline">NetOffice.MSProjectApi.Enums.PjBaselines baseline</param>
		[SupportByVersion("MSProject", 11,12,14), Redirect("get_BaselineSavedDate")]
		public object BaselineSavedDate(NetOffice.MSProjectApi.Enums.PjBaselines baseline)
		{
			return get_BaselineSavedDate(baseline);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string ProjectNamePrefix
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ProjectNamePrefix");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string VersionName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "VersionName");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 TempToDoList
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TempToDoList");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TempToDoList", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.OutlineCodes OutlineCodes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.OutlineCodes>(this, "OutlineCodes", NetOffice.MSProjectApi.OutlineCodes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.OfficeApi.SharedWorkspace SharedWorkspace
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SharedWorkspace>(this, "SharedWorkspace", NetOffice.OfficeApi.SharedWorkspace.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.OfficeApi.DocumentLibraryVersions DocumentLibraryVersions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DocumentLibraryVersions>(this, "DocumentLibraryVersions", NetOffice.OfficeApi.DocumentLibraryVersions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool CanCheckIn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CanCheckIn");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string CurrencyCode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CurrencyCode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CurrencyCode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 TaskErrorCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TaskErrorCount");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TaskErrorCount", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool IsTemplate
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsTemplate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IsTemplate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public Int32 HyperlinkColorEx
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HyperlinkColorEx");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HyperlinkColorEx", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public Int32 FollowedHyperlinkColorEx
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "FollowedHyperlinkColorEx");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FollowedHyperlinkColorEx", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public bool NewTasksCreatedAsManual
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "NewTasksCreatedAsManual");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NewTasksCreatedAsManual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public NetOffice.MSProjectApi.TaskGroups2 TaskGroups2
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.TaskGroups2>(this, "TaskGroups2", NetOffice.MSProjectApi.TaskGroups2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public NetOffice.MSProjectApi.ResourceGroups2 ResourceGroups2
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.ResourceGroups2>(this, "ResourceGroups2", NetOffice.MSProjectApi.ResourceGroups2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public bool ManuallyScheduledTasksAutoRespectLinks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ManuallyScheduledTasksAutoRespectLinks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ManuallyScheduledTasksAutoRespectLinks", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public bool KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public bool ShowTaskWarnings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowTaskWarnings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowTaskWarnings", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public bool ShowTaskSuggestions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowTaskSuggestions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowTaskSuggestions", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public NetOffice.MSProjectApi.Tasks DetectCycle
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Tasks>(this, "DetectCycle", NetOffice.MSProjectApi.Tasks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public NetOffice.MSProjectApi.Reports Reports
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Reports>(this, "Reports", NetOffice.MSProjectApi.Reports.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public bool IsCheckoutMsgBarVisible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsCheckoutMsgBarVisible");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public bool IsCheckoutOSVisible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsCheckoutOSVisible");
			}
		}

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
		public void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord, object formatID, object map, object clearBaseline, object clearActuals, object clearResourceRates, object clearFixedCosts)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered, table, userID, databasePassWord, formatID, map, clearBaseline, clearActuals, clearResourceRates, clearFixedCosts });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public void SaveAs(object name)
		{
			 Factory.ExecuteMethod(this, "SaveAs", name);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public void SaveAs(object name, object format)
		{
			 Factory.ExecuteMethod(this, "SaveAs", name, format);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		/// <param name="backup">optional object backup</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public void SaveAs(object name, object format, object backup)
		{
			 Factory.ExecuteMethod(this, "SaveAs", name, format, backup);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		/// <param name="backup">optional object backup</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public void SaveAs(object name, object format, object backup, object readOnly)
		{
			 Factory.ExecuteMethod(this, "SaveAs", name, format, backup, readOnly);
		}

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
		public void SaveAs(object name, object format, object backup, object readOnly, object taskInformation)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation });
		}

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
		public void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered });
		}

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
		public void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered, table });
		}

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
		public void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered, table, userID });
		}

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
		public void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered, table, userID, databasePassWord });
		}

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
		public void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord, object formatID)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered, table, userID, databasePassWord, formatID });
		}

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
		public void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord, object formatID, object map)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered, table, userID, databasePassWord, formatID, map });
		}

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
		public void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord, object formatID, object map, object clearBaseline)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered, table, userID, databasePassWord, formatID, map, clearBaseline });
		}

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
		public void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord, object formatID, object map, object clearBaseline, object clearActuals)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered, table, userID, databasePassWord, formatID, map, clearBaseline, clearActuals });
		}

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
		public void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord, object formatID, object map, object clearBaseline, object clearActuals, object clearResourceRates)
		{
			 Factory.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered, table, userID, databasePassWord, formatID, map, clearBaseline, clearActuals, clearResourceRates });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public void Activate()
		{
			 Factory.ExecuteMethod(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public void LevelClearDates()
		{
			 Factory.ExecuteMethod(this, "LevelClearDates");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="value">string value</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void AppendNotes(string value)
		{
			 Factory.ExecuteMethod(this, "AppendNotes", value);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public void MakeServerURLTrusted()
		{
			 Factory.ExecuteMethod(this, "MakeServerURLTrusted");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="comment">optional object comment</param>
		/// <param name="makePublic">optional object makePublic</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void CheckIn(object saveChanges, object comment, object makePublic)
		{
			 Factory.ExecuteMethod(this, "CheckIn", saveChanges, comment, makePublic);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public void CheckIn()
		{
			 Factory.ExecuteMethod(this, "CheckIn");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="saveChanges">optional object saveChanges</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public void CheckIn(object saveChanges)
		{
			 Factory.ExecuteMethod(this, "CheckIn", saveChanges);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="comment">optional object comment</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public void CheckIn(object saveChanges, object comment)
		{
			 Factory.ExecuteMethod(this, "CheckIn", saveChanges, comment);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="objectType">NetOffice.MSProjectApi.Enums.PjOrganizer objectType</param>
		/// <param name="objectName">string objectName</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public string GetObjectMatchingID(NetOffice.MSProjectApi.Enums.PjOrganizer objectType, string objectName)
		{
			return Factory.ExecuteStringMethodGet(this, "GetObjectMatchingID", objectType, objectName);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="objectType">NetOffice.MSProjectApi.Enums.PjOrganizer objectType</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="matchingID">string matchingID</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void SetObjectMatchingID(NetOffice.MSProjectApi.Enums.PjOrganizer objectType, string objectName, string matchingID)
		{
			 Factory.ExecuteMethod(this, "SetObjectMatchingID", objectType, objectName, matchingID);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="objectType">NetOffice.MSProjectApi.Enums.PjOrganizer objectType</param>
		/// <param name="matchingID">string matchingID</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public string GetDisplayNameFromObjectMatchingID(NetOffice.MSProjectApi.Enums.PjOrganizer objectType, string matchingID)
		{
			return Factory.ExecuteStringMethodGet(this, "GetDisplayNameFromObjectMatchingID", objectType, matchingID);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableName">string deliverableName</param>
		/// <param name="deliverableStartDate">object deliverableStartDate</param>
		/// <param name="deliverableFinishDate">object deliverableFinishDate</param>
		/// <param name="taskGuid">string taskGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public string DeliverableCreate(string deliverableName, object deliverableStartDate, object deliverableFinishDate, string taskGuid)
		{
			return Factory.ExecuteStringMethodGet(this, "DeliverableCreate", deliverableName, deliverableStartDate, deliverableFinishDate, taskGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		/// <param name="deliverableName">string deliverableName</param>
		/// <param name="deliverableStartDate">object deliverableStartDate</param>
		/// <param name="deliverableFinishDate">object deliverableFinishDate</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool DeliverableUpdate(string deliverableGuid, string deliverableName, object deliverableStartDate, object deliverableFinishDate)
		{
			return Factory.ExecuteBoolMethodGet(this, "DeliverableUpdate", deliverableGuid, deliverableName, deliverableStartDate, deliverableFinishDate);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool DeliverableDelete(string deliverableGuid)
		{
			return Factory.ExecuteBoolMethodGet(this, "DeliverableDelete", deliverableGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		/// <param name="taskGuid">string taskGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool DeliverableDependencyCreate(string deliverableGuid, string taskGuid)
		{
			return Factory.ExecuteBoolMethodGet(this, "DeliverableDependencyCreate", deliverableGuid, taskGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool DeliverableDependencyDelete(string deliverableGuid)
		{
			return Factory.ExecuteBoolMethodGet(this, "DeliverableDependencyDelete", deliverableGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">optional object deliverableGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool DeliverableRefreshServerCache(object deliverableGuid)
		{
			return Factory.ExecuteBoolMethodGet(this, "DeliverableRefreshServerCache", deliverableGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public bool DeliverableRefreshServerCache()
		{
			return Factory.ExecuteBoolMethodGet(this, "DeliverableRefreshServerCache");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object DeliverablesGetServerCachedXml()
		{
			return Factory.ExecuteVariantMethodGet(this, "DeliverablesGetServerCachedXml");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object DeliverablesGetXml()
		{
			return Factory.ExecuteVariantMethodGet(this, "DeliverablesGetXml");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string GetServerProjectGuid()
		{
			return Factory.ExecuteStringMethodGet(this, "GetServerProjectGuid");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		/// <param name="taskGuid">string taskGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool DeliverableLinkToTask(string deliverableGuid, string taskGuid)
		{
			return Factory.ExecuteBoolMethodGet(this, "DeliverableLinkToTask", deliverableGuid, taskGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool DeliverableLinkToProject(string deliverableGuid)
		{
			return Factory.ExecuteBoolMethodGet(this, "DeliverableLinkToProject", deliverableGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool DeliverablesClearAll()
		{
			return Factory.ExecuteBoolMethodGet(this, "DeliverablesClearAll");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool DeliverableAcceptChanges(string deliverableGuid)
		{
			return Factory.ExecuteBoolMethodGet(this, "DeliverableAcceptChanges", deliverableGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string DeliverablesGetProviderProjects()
		{
			return Factory.ExecuteStringMethodGet(this, "DeliverablesGetProviderProjects");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="projectGuid">string projectGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public object DeliverablesGetByProject(string projectGuid)
		{
			return Factory.ExecuteVariantMethodGet(this, "DeliverablesGetByProject", projectGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="taskGuid">string taskGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 GetTaskIndexByGuid(string taskGuid)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetTaskIndexByGuid", taskGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="projectGuid">string projectGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public object ReadWssData(string projectGuid)
		{
			return Factory.ExecuteVariantMethodGet(this, "ReadWssData", projectGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object GetWinprojURLs()
		{
			return Factory.ExecuteVariantMethodGet(this, "GetWinprojURLs");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 LocalResourceErrorCount()
		{
			return Factory.ExecuteInt32MethodGet(this, "LocalResourceErrorCount");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 ImportResourceErrorCount()
		{
			return Factory.ExecuteInt32MethodGet(this, "ImportResourceErrorCount");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 ResourceErrorCount()
		{
			return Factory.ExecuteInt32MethodGet(this, "ResourceErrorCount");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 LocalResourceCount()
		{
			return Factory.ExecuteInt32MethodGet(this, "LocalResourceCount");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 ResourceCount()
		{
			return Factory.ExecuteInt32MethodGet(this, "ResourceCount");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="assignmentToDrag">object assignmentToDrag</param>
		/// <param name="destinationResource">optional object destinationResource</param>
		/// <param name="destinationTime">optional object destinationTime</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSProject", 11,14)]
		public void RSVDragSimulator(object assignmentToDrag, object destinationResource, object destinationTime)
		{
			 Factory.ExecuteMethod(this, "RSVDragSimulator", assignmentToDrag, destinationResource, destinationTime);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="assignmentToDrag">object assignmentToDrag</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public void RSVDragSimulator(object assignmentToDrag)
		{
			 Factory.ExecuteMethod(this, "RSVDragSimulator", assignmentToDrag);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="assignmentToDrag">object assignmentToDrag</param>
		/// <param name="destinationResource">optional object destinationResource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public void RSVDragSimulator(object assignmentToDrag, object destinationResource)
		{
			 Factory.ExecuteMethod(this, "RSVDragSimulator", assignmentToDrag, destinationResource);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="customUIXML">string customUIXML</param>
		[SupportByVersion("MSProject", 11,14)]
		public void SetCustomUI(string customUIXML)
		{
			 Factory.ExecuteMethod(this, "SetCustomUI", customUIXML);
		}

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
		public void ExportAsFixedFormat(string filename, object fileType, object includeDocumentProperties, object includeDocumentMarkup, object archiveFormat, object fromDate, object toDate, object fixedFormatExtClassPtr)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ filename, fileType, includeDocumentProperties, includeDocumentMarkup, archiveFormat, fromDate, toDate, fixedFormatExtClassPtr });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public void ExportAsFixedFormat(string filename)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", filename);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileType">optional NetOffice.MSProjectApi.Enums.PjDocExportType FileType = 0</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public void ExportAsFixedFormat(string filename, object fileType)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", filename, fileType);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileType">optional NetOffice.MSProjectApi.Enums.PjDocExportType FileType = 0</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public void ExportAsFixedFormat(string filename, object fileType, object includeDocumentProperties)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", filename, fileType, includeDocumentProperties);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileType">optional NetOffice.MSProjectApi.Enums.PjDocExportType FileType = 0</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="includeDocumentMarkup">optional bool IncludeDocumentMarkup = true</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public void ExportAsFixedFormat(string filename, object fileType, object includeDocumentProperties, object includeDocumentMarkup)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", filename, fileType, includeDocumentProperties, includeDocumentMarkup);
		}

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
		public void ExportAsFixedFormat(string filename, object fileType, object includeDocumentProperties, object includeDocumentMarkup, object archiveFormat)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ filename, fileType, includeDocumentProperties, includeDocumentMarkup, archiveFormat });
		}

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
		public void ExportAsFixedFormat(string filename, object fileType, object includeDocumentProperties, object includeDocumentMarkup, object archiveFormat, object fromDate)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ filename, fileType, includeDocumentProperties, includeDocumentMarkup, archiveFormat, fromDate });
		}

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
		public void ExportAsFixedFormat(string filename, object fileType, object includeDocumentProperties, object includeDocumentMarkup, object archiveFormat, object fromDate, object toDate)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ filename, fileType, includeDocumentProperties, includeDocumentMarkup, archiveFormat, fromDate, toDate });
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public Int32 CheckoutProject()
		{
			return Factory.ExecuteInt32MethodGet(this, "CheckoutProject");
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public Int32 HideCheckoutMsgBar()
		{
			return Factory.ExecuteInt32MethodGet(this, "HideCheckoutMsgBar");
		}

		#endregion

		#pragma warning restore
	}
}
