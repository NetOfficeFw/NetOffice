using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSProjectApi;

namespace NetOffice.MSProjectApi.Behind
{
	/// <summary>
	/// DispatchInterface _IProjectDoc 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _IProjectDoc : COMObject, NetOffice.MSProjectApi._IProjectDoc
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
                    _contractType = typeof(NetOffice.MSProjectApi._IProjectDoc);
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
                    _type = typeof(_IProjectDoc);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _IProjectDoc() : base()
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
		public virtual object Manager
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Manager");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Manager", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Company
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Company");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Company", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Author
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Author");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Author", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Keywords
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Keywords");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Keywords", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string ProjectNotes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ProjectNotes");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ProjectNotes", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object ProjectStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ProjectStart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ProjectStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object ProjectFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ProjectFinish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ProjectFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object CurrentDate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "CurrentDate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "CurrentDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object StatusDate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "StatusDate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "StatusDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool ScheduleFromStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ScheduleFromStart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScheduleFromStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Comments
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Comments");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Comments", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Title
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Title");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Title", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Subject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Subject");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Subject", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Windows Windows
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Windows>(this, "Windows", typeof(NetOffice.MSProjectApi.Windows));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Windows", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int16 MinuteLabelDisplay
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "MinuteLabelDisplay");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MinuteLabelDisplay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int16 HourLabelDisplay
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "HourLabelDisplay");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HourLabelDisplay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int16 DayLabelDisplay
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "DayLabelDisplay");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DayLabelDisplay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int16 WeekLabelDisplay
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "WeekLabelDisplay");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WeekLabelDisplay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int16 YearLabelDisplay
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "YearLabelDisplay");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "YearLabelDisplay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int16 MonthLabelDisplay
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "MonthLabelDisplay");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MonthLabelDisplay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool SpaceBeforeTimeLabels
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SpaceBeforeTimeLabels");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SpaceBeforeTimeLabels", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjTaskFixedType DefaultTaskType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjTaskFixedType>(this, "DefaultTaskType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DefaultTaskType", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool DefaultEffortDriven
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DefaultEffortDriven");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultEffortDriven", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool UseFYStartYear
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UseFYStartYear");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UseFYStartYear", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool AutoFilter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoFilter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool HonorConstraints
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HonorConstraints");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HonorConstraints", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool MultipleCriticalPaths
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MultipleCriticalPaths");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MultipleCriticalPaths", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object LevelFromDate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "LevelFromDate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "LevelFromDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object LevelToDate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "LevelToDate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "LevelToDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool LevelEntireProject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "LevelEntireProject");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LevelEntireProject", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjAccrueAt DefaultFixedCostAccrual
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "DefaultFixedCostAccrual");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DefaultFixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool SpreadCostsToStatusDate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SpreadCostsToStatusDate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SpreadCostsToStatusDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool SpreadPercentCompleteToStatusDate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SpreadPercentCompleteToStatusDate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SpreadPercentCompleteToStatusDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool AutoCalcCosts
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoCalcCosts");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoCalcCosts", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool ShowExternalSuccessors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowExternalSuccessors");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowExternalSuccessors", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool ShowExternalPredecessors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowExternalPredecessors");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowExternalPredecessors", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool ShowCrossProjectLinksInfo
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowCrossProjectLinksInfo");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowCrossProjectLinksInfo", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool AcceptNewExternalData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AcceptNewExternalData");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AcceptNewExternalData", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjPhoneticType PhoneticType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjPhoneticType>(this, "PhoneticType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PhoneticType", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjWorkgroupMessages WorkgroupMessages
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjWorkgroupMessages>(this, "WorkgroupMessages");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "WorkgroupMessages", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string ServerURL
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ServerURL");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ServerURL", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string ServerPath
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ServerPath");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ServerPath", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool ReceiveNotifications
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ReceiveNotifications");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ReceiveNotifications", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool SendHyperlinkNote
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SendHyperlinkNote");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SendHyperlinkNote", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjColor HyperlinkColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjColor>(this, "HyperlinkColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "HyperlinkColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjColor FollowedHyperlinkColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjColor>(this, "FollowedHyperlinkColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "FollowedHyperlinkColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool UnderlineHyperlinks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UnderlineHyperlinks");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UnderlineHyperlinks", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjTeamStatusCompletedWork AskForCompletedWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjTeamStatusCompletedWork>(this, "AskForCompletedWork");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "AskForCompletedWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool TrackOvertimeWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "TrackOvertimeWork");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TrackOvertimeWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool TeamMembersCanDeclineTasks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "TeamMembersCanDeclineTasks");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TeamMembersCanDeclineTasks", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool ShowEstimatedDuration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowEstimatedDuration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowEstimatedDuration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool NewTasksEstimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "NewTasksEstimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NewTasksEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool WBSCodeGenerate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "WBSCodeGenerate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WBSCodeGenerate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool WBSVerifyUniqueness
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "WBSVerifyUniqueness");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WBSVerifyUniqueness", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool UpdateProjOnSave
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UpdateProjOnSave");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UpdateProjOnSave", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjAuthentication ServerIdentification
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAuthentication>(this, "ServerIdentification");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ServerIdentification", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool VBASigned
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "VBASigned");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "VBASigned", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool ExpandDatabaseTimephasedData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ExpandDatabaseTimephasedData");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ExpandDatabaseTimephasedData", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object DatabaseProjectUniqueID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DatabaseProjectUniqueID");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "DatabaseProjectUniqueID", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object ActualWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ActualWork");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ActualWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Cost1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Cost1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Cost1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Cost2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Cost2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Cost2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Cost3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Cost3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Cost3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object BaselineWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BaselineWork");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BaselineWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object BaselineCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BaselineCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BaselineCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object FixedCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FixedCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string WBS
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "WBS");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WBS", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Delay
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Delay");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Delay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Priority
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Priority");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Priority", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Duration1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Duration2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Duration3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object PercentWorkComplete
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PercentWorkComplete");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "PercentWorkComplete", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object FixedDuration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FixedDuration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "FixedDuration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object BaselineStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BaselineStart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BaselineStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object BaselineFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BaselineFinish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BaselineFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string Text1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Start1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Start1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Start1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Finish1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Finish1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Finish1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string Text2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Start2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Start2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Start2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Finish2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Finish2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Finish2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string Text3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Start3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Start3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Start3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Finish3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Finish3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Finish3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string Text4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Start4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Start4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Start4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Finish4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Finish4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Finish4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string Text5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Start5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Start5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Start5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Finish5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Finish5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Finish5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string Text6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string Text7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string Text8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string Text9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string Text10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Marked
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Marked");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Marked", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Flag1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Flag2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Flag3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Flag4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Flag5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Flag6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Flag7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Flag8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Flag9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Flag10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Rollup
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Rollup");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Rollup", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Double Number1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Double Number2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Double Number3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Double Number4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Double Number5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string Notes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Notes");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Notes", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string Contact
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Contact");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Contact", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object HideBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "HideBar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "HideBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string CurrencySymbol
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CurrencySymbol");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CurrencySymbol", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjPlacement CurrencySymbolPosition
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjPlacement>(this, "CurrencySymbolPosition");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "CurrencySymbolPosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int16 CurrencyDigits
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "CurrencyDigits");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CurrencyDigits", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int32 ShowCriticalSlack
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ShowCriticalSlack");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowCriticalSlack", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjUnit DefaultDurationUnits
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjUnit>(this, "DefaultDurationUnits");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DefaultDurationUnits", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjUnit DefaultWorkUnits
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjUnit>(this, "DefaultWorkUnits");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DefaultWorkUnits", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool StartOnCurrentDate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "StartOnCurrentDate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StartOnCurrentDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool AutoTrack
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoTrack");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoTrack", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool AutoSplitTasks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoSplitTasks");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoSplitTasks", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool AutoLinkTasks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoLinkTasks");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoLinkTasks", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object DefaultStartTime
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DefaultStartTime");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "DefaultStartTime", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object DefaultFinishTime
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DefaultFinishTime");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "DefaultFinishTime", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double HoursPerDay
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "HoursPerDay");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HoursPerDay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double HoursPerWeek
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "HoursPerWeek");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HoursPerWeek", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double DaysPerMonth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "DaysPerMonth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DaysPerMonth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object DefaultResourceStandardRate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DefaultResourceStandardRate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "DefaultResourceStandardRate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object DefaultResourceOvertimeRate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DefaultResourceOvertimeRate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "DefaultResourceOvertimeRate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool DisplayProjectSummaryTask
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayProjectSummaryTask");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayProjectSummaryTask", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool AutoAddResources
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoAddResources");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoAddResources", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjWeekday StartWeekOn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjWeekday>(this, "StartWeekOn");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "StartWeekOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjMonth StartYearIn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjMonth>(this, "StartYearIn");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "StartYearIn", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool AllowTaskDelegation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowTaskDelegation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowTaskDelegation", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjPublishInformationOnSave PublishInformationOnSave
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjPublishInformationOnSave>(this, "PublishInformationOnSave");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PublishInformationOnSave", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string ProjectGuideFunctionalLayoutPage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ProjectGuideFunctionalLayoutPage");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ProjectGuideFunctionalLayoutPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string ProjectGuideSaveBuffer
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ProjectGuideSaveBuffer");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ProjectGuideSaveBuffer", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string ProjectGuideContent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ProjectGuideContent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ProjectGuideContent", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool ProjectServerUsedForTracking
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ProjectServerUsedForTracking");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ProjectServerUsedForTracking", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjProjectServerTrackingMethod TrackingMethod
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjProjectServerTrackingMethod>(this, "TrackingMethod");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "TrackingMethod", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool MoveCompleted
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MoveCompleted");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MoveCompleted", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool AndMoveRemaining
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AndMoveRemaining");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AndMoveRemaining", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool MoveRemaining
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MoveRemaining");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MoveRemaining", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool AndMoveCompleted
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AndMoveCompleted");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AndMoveCompleted", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjEarnedValueMethod DefaultEarnedValueMethod
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjEarnedValueMethod>(this, "DefaultEarnedValueMethod");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DefaultEarnedValueMethod", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjBaselines EarnedValueBaseline
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjBaselines>(this, "EarnedValueBaseline");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "EarnedValueBaseline", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool ProjectGuideUseDefaultFunctionalLayoutPage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ProjectGuideUseDefaultFunctionalLayoutPage");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ProjectGuideUseDefaultFunctionalLayoutPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool ProjectGuideUseDefaultContent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ProjectGuideUseDefaultContent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ProjectGuideUseDefaultContent", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool EnterpriseActualsSynched
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnterpriseActualsSynched");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseActualsSynched", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool RemoveFileProperties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "RemoveFileProperties");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RemoveFileProperties", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool AdministrativeProject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AdministrativeProject");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AdministrativeProject", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Windows2 Windows2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Windows2>(this, "Windows2", typeof(NetOffice.MSProjectApi.Windows2));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Windows2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string _CodeName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "_CodeName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "_CodeName", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string CodeName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CodeName");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Tasks OutlineChildren
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Tasks>(this, "OutlineChildren", typeof(NetOffice.MSProjectApi.Tasks));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object CostVariance
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "CostVariance");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Task ProjectSummaryTask
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Task>(this, "ProjectSummaryTask", typeof(NetOffice.MSProjectApi.Task));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object RemainingCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "RemainingCost");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object BCWP
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BCWP");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object BCWS
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BCWS");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object SV
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "SV");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object CV
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "CV");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OutlineNumber
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OutlineNumber");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Critical
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Critical");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object FreeSlack
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FreeSlack");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object TotalSlack
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "TotalSlack");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int32 UniqueID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "UniqueID");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 OutlineLevel
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "OutlineLevel");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object BaselineDuration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BaselineDuration");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object DurationVariance
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DurationVariance");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object EarlyStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EarlyStart");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object EarlyFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EarlyFinish");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object LateStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "LateStart");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object StartVariance
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "StartVariance");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object FinishVariance
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FinishVariance");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Project
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Project");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Milestone
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Milestone");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object RemainingDuration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "RemainingDuration");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object PercentComplete
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PercentComplete");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Start
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Start");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Finish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Finish");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object ResourceNames
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ResourceNames");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object ResourceInitials
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ResourceInitials");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Resume
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Resume");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Stop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Stop");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object ResumeNoEarlierThan
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ResumeNoEarlierThan");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object ConstraintType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ConstraintType");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object ConstraintDate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ConstraintDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object ActualCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ActualCost");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Cost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Cost");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Created
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Created");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object ActualDuration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ActualDuration");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Duration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object LateFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "LateFinish");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object ActualFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ActualFinish");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 Objects
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Objects");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object RemainingWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "RemainingWork");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object ResourceGroup
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ResourceGroup");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object ActualStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ActualStart");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Summary
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Summary");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Template
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Template");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object UpdateNeeded
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "UpdateNeeded");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Work
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Work");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object WorkVariance
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "WorkVariance");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object LinkedFields
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "LinkedFields");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object Confirmed
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Confirmed");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool ReadOnly
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ReadOnly");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool HasPassword
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasPassword");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool WriteReserved
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "WriteReserved");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Application>(this, "Application", typeof(NetOffice.MSProjectApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Index
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.List MapList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "MapList", typeof(NetOffice.MSProjectApi.List));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Tasks Tasks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Tasks>(this, "Tasks", typeof(NetOffice.MSProjectApi.Tasks));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Resources Resources
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Resources>(this, "Resources", typeof(NetOffice.MSProjectApi.Resources));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Calendars BaseCalendars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Calendars>(this, "BaseCalendars", typeof(NetOffice.MSProjectApi.Calendars));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14), ProxyResult]
		public virtual object BuiltinDocumentProperties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "BuiltinDocumentProperties");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14), ProxyResult]
		public virtual object CustomDocumentProperties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "CustomDocumentProperties");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14), ProxyResult]
		public virtual object Container
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Container");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Calendar Calendar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Calendar>(this, "Calendar", typeof(NetOffice.MSProjectApi.Calendar));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int32 NumberOfTasks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "NumberOfTasks");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int32 NumberOfResources
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "NumberOfResources");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string FullName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FullName");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Path
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Path");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string ResourcePoolName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ResourcePoolName");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool Saved
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Saved");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object CreationDate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "CreationDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object LastPrintedDate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "LastPrintedDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object LastSaveDate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "LastSaveDate");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string LastSavedBy
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "LastSavedBy");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string RevisionNumber
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RevisionNumber");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.List ViewList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "ViewList", typeof(NetOffice.MSProjectApi.List));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.List TaskViewList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "TaskViewList", typeof(NetOffice.MSProjectApi.List));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.List ResourceViewList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "ResourceViewList", typeof(NetOffice.MSProjectApi.List));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool ReadOnlyRecommended
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ReadOnlyRecommended");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.List ReportList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "ReportList", typeof(NetOffice.MSProjectApi.List));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.List TaskFilterList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "TaskFilterList", typeof(NetOffice.MSProjectApi.List));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.List ResourceFilterList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "ResourceFilterList", typeof(NetOffice.MSProjectApi.List));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.List TaskTableList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "TaskTableList", typeof(NetOffice.MSProjectApi.List));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.List ResourceTableList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "ResourceTableList", typeof(NetOffice.MSProjectApi.List));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string CurrentView
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CurrentView");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string CurrentTable
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CurrentTable");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string CurrentFilter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CurrentFilter");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int32 ID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ID");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.OfficeApi.CommandBars CommandBars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(this, "CommandBars", typeof(NetOffice.OfficeApi.CommandBars));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool UserControl
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UserControl");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.VBIDEApi.VBProject VBProject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBProject>(this, "VBProject", typeof(NetOffice.VBIDEApi.VBProject));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Subprojects Subprojects
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Subprojects>(this, "Subprojects", typeof(NetOffice.MSProjectApi.Subprojects));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string CurrentGroup
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CurrentGroup");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.List TaskGroupList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "TaskGroupList", typeof(NetOffice.MSProjectApi.List));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.List ResourceGroupList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.List>(this, "ResourceGroupList", typeof(NetOffice.MSProjectApi.List));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.TaskGroups TaskGroups
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.TaskGroups>(this, "TaskGroups", typeof(NetOffice.MSProjectApi.TaskGroups));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.ResourceGroups ResourceGroups
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.ResourceGroups>(this, "ResourceGroups", typeof(NetOffice.MSProjectApi.ResourceGroups));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjProjectType Type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjProjectType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[BaseResult]
		public virtual NetOffice.MSProjectApi.Views Views
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.MSProjectApi.Views>(this, "Views");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Tables TaskTables
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Tables>(this, "TaskTables", typeof(NetOffice.MSProjectApi.Tables));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Tables ResourceTables
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Tables>(this, "ResourceTables", typeof(NetOffice.MSProjectApi.Tables));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Filters TaskFilters
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Filters>(this, "TaskFilters", typeof(NetOffice.MSProjectApi.Filters));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Filters ResourceFilters
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Filters>(this, "ResourceFilters", typeof(NetOffice.MSProjectApi.Filters));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.ViewsSingle ViewsSingle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.ViewsSingle>(this, "ViewsSingle", typeof(NetOffice.MSProjectApi.ViewsSingle));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.ViewsCombination ViewsCombination
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.ViewsCombination>(this, "ViewsCombination", typeof(NetOffice.MSProjectApi.ViewsCombination));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		/// <param name="baseline">NetOffice.MSProjectApi.Enums.PjBaselines baseline</param>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object get_BaselineSavedDate(NetOffice.MSProjectApi.Enums.PjBaselines baseline)
		{
			return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BaselineSavedDate", baseline);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Alias for get_BaselineSavedDate
		/// </summary>
		/// <param name="baseline">NetOffice.MSProjectApi.Enums.PjBaselines baseline</param>
		[SupportByVersion("MSProject", 11,12,14), Redirect("get_BaselineSavedDate")]
		public virtual object BaselineSavedDate(NetOffice.MSProjectApi.Enums.PjBaselines baseline)
		{
			return get_BaselineSavedDate(baseline);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string ProjectNamePrefix
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ProjectNamePrefix");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string VersionName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "VersionName");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 TempToDoList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "TempToDoList");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TempToDoList", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.OutlineCodes OutlineCodes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.OutlineCodes>(this, "OutlineCodes", typeof(NetOffice.MSProjectApi.OutlineCodes));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.OfficeApi.SharedWorkspace SharedWorkspace
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SharedWorkspace>(this, "SharedWorkspace", typeof(NetOffice.OfficeApi.SharedWorkspace));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.OfficeApi.DocumentLibraryVersions DocumentLibraryVersions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DocumentLibraryVersions>(this, "DocumentLibraryVersions", typeof(NetOffice.OfficeApi.DocumentLibraryVersions));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool CanCheckIn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CanCheckIn");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string CurrencyCode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CurrencyCode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CurrencyCode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int32 TaskErrorCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "TaskErrorCount");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TaskErrorCount", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool IsTemplate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsTemplate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IsTemplate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual Int32 HyperlinkColorEx
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "HyperlinkColorEx");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HyperlinkColorEx", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual Int32 FollowedHyperlinkColorEx
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "FollowedHyperlinkColorEx");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FollowedHyperlinkColorEx", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual bool NewTasksCreatedAsManual
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "NewTasksCreatedAsManual");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NewTasksCreatedAsManual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.TaskGroups2 TaskGroups2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.TaskGroups2>(this, "TaskGroups2", typeof(NetOffice.MSProjectApi.TaskGroups2));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.ResourceGroups2 ResourceGroups2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.ResourceGroups2>(this, "ResourceGroups2", typeof(NetOffice.MSProjectApi.ResourceGroups2));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual bool ManuallyScheduledTasksAutoRespectLinks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ManuallyScheduledTasksAutoRespectLinks");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ManuallyScheduledTasksAutoRespectLinks", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual bool KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual bool ShowTaskWarnings
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowTaskWarnings");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowTaskWarnings", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual bool ShowTaskSuggestions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowTaskSuggestions");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowTaskSuggestions", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.Tasks DetectCycle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Tasks>(this, "DetectCycle", typeof(NetOffice.MSProjectApi.Tasks));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public virtual NetOffice.MSProjectApi.Reports Reports
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Reports>(this, "Reports", typeof(NetOffice.MSProjectApi.Reports));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public virtual bool IsCheckoutMsgBarVisible
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsCheckoutMsgBarVisible");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public virtual bool IsCheckoutOSVisible
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsCheckoutOSVisible");
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
		public virtual void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord, object formatID, object map, object clearBaseline, object clearActuals, object clearResourceRates, object clearFixedCosts)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered, table, userID, databasePassWord, formatID, map, clearBaseline, clearActuals, clearResourceRates, clearFixedCosts });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void SaveAs(object name)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", name);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void SaveAs(object name, object format)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", name, format);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">object name</param>
		/// <param name="format">optional NetOffice.MSProjectApi.Enums.PjFileFormat Format = 0</param>
		/// <param name="backup">optional object backup</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void SaveAs(object name, object format, object backup)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", name, format, backup);
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
		public virtual void SaveAs(object name, object format, object backup, object readOnly)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", name, format, backup, readOnly);
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
		public virtual void SaveAs(object name, object format, object backup, object readOnly, object taskInformation)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation });
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
		public virtual void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered });
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
		public virtual void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered, table });
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
		public virtual void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered, table, userID });
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
		public virtual void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered, table, userID, databasePassWord });
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
		public virtual void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord, object formatID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered, table, userID, databasePassWord, formatID });
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
		public virtual void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord, object formatID, object map)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered, table, userID, databasePassWord, formatID, map });
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
		public virtual void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord, object formatID, object map, object clearBaseline)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered, table, userID, databasePassWord, formatID, map, clearBaseline });
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
		public virtual void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord, object formatID, object map, object clearBaseline, object clearActuals)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered, table, userID, databasePassWord, formatID, map, clearBaseline, clearActuals });
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
		public virtual void SaveAs(object name, object format, object backup, object readOnly, object taskInformation, object filtered, object table, object userID, object databasePassWord, object formatID, object map, object clearBaseline, object clearActuals, object clearResourceRates)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[]{ name, format, backup, readOnly, taskInformation, filtered, table, userID, databasePassWord, formatID, map, clearBaseline, clearActuals, clearResourceRates });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void Activate()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void LevelClearDates()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LevelClearDates");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="value">string value</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void AppendNotes(string value)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AppendNotes", value);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void MakeServerURLTrusted()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MakeServerURLTrusted");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="comment">optional object comment</param>
		/// <param name="makePublic">optional object makePublic</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void CheckIn(object saveChanges, object comment, object makePublic)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn", saveChanges, comment, makePublic);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void CheckIn()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="saveChanges">optional object saveChanges</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void CheckIn(object saveChanges)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn", saveChanges);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="comment">optional object comment</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void CheckIn(object saveChanges, object comment)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn", saveChanges, comment);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="objectType">NetOffice.MSProjectApi.Enums.PjOrganizer objectType</param>
		/// <param name="objectName">string objectName</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string GetObjectMatchingID(NetOffice.MSProjectApi.Enums.PjOrganizer objectType, string objectName)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetObjectMatchingID", objectType, objectName);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="objectType">NetOffice.MSProjectApi.Enums.PjOrganizer objectType</param>
		/// <param name="objectName">string objectName</param>
		/// <param name="matchingID">string matchingID</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void SetObjectMatchingID(NetOffice.MSProjectApi.Enums.PjOrganizer objectType, string objectName, string matchingID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetObjectMatchingID", objectType, objectName, matchingID);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="objectType">NetOffice.MSProjectApi.Enums.PjOrganizer objectType</param>
		/// <param name="matchingID">string matchingID</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string GetDisplayNameFromObjectMatchingID(NetOffice.MSProjectApi.Enums.PjOrganizer objectType, string matchingID)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetDisplayNameFromObjectMatchingID", objectType, matchingID);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableName">string deliverableName</param>
		/// <param name="deliverableStartDate">object deliverableStartDate</param>
		/// <param name="deliverableFinishDate">object deliverableFinishDate</param>
		/// <param name="taskGuid">string taskGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string DeliverableCreate(string deliverableName, object deliverableStartDate, object deliverableFinishDate, string taskGuid)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "DeliverableCreate", deliverableName, deliverableStartDate, deliverableFinishDate, taskGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		/// <param name="deliverableName">string deliverableName</param>
		/// <param name="deliverableStartDate">object deliverableStartDate</param>
		/// <param name="deliverableFinishDate">object deliverableFinishDate</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool DeliverableUpdate(string deliverableGuid, string deliverableName, object deliverableStartDate, object deliverableFinishDate)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "DeliverableUpdate", deliverableGuid, deliverableName, deliverableStartDate, deliverableFinishDate);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool DeliverableDelete(string deliverableGuid)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "DeliverableDelete", deliverableGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		/// <param name="taskGuid">string taskGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool DeliverableDependencyCreate(string deliverableGuid, string taskGuid)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "DeliverableDependencyCreate", deliverableGuid, taskGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool DeliverableDependencyDelete(string deliverableGuid)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "DeliverableDependencyDelete", deliverableGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">optional object deliverableGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool DeliverableRefreshServerCache(object deliverableGuid)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "DeliverableRefreshServerCache", deliverableGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool DeliverableRefreshServerCache()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "DeliverableRefreshServerCache");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object DeliverablesGetServerCachedXml()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DeliverablesGetServerCachedXml");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object DeliverablesGetXml()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DeliverablesGetXml");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string GetServerProjectGuid()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetServerProjectGuid");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		/// <param name="taskGuid">string taskGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool DeliverableLinkToTask(string deliverableGuid, string taskGuid)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "DeliverableLinkToTask", deliverableGuid, taskGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool DeliverableLinkToProject(string deliverableGuid)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "DeliverableLinkToProject", deliverableGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool DeliverablesClearAll()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "DeliverablesClearAll");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="deliverableGuid">string deliverableGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual bool DeliverableAcceptChanges(string deliverableGuid)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "DeliverableAcceptChanges", deliverableGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string DeliverablesGetProviderProjects()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "DeliverablesGetProviderProjects");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="projectGuid">string projectGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object DeliverablesGetByProject(string projectGuid)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DeliverablesGetByProject", projectGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="taskGuid">string taskGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int32 GetTaskIndexByGuid(string taskGuid)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetTaskIndexByGuid", taskGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="projectGuid">string projectGuid</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object ReadWssData(string projectGuid)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ReadWssData", projectGuid);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object GetWinprojURLs()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetWinprojURLs");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int32 LocalResourceErrorCount()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "LocalResourceErrorCount");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int32 ImportResourceErrorCount()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ImportResourceErrorCount");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int32 ResourceErrorCount()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ResourceErrorCount");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int32 LocalResourceCount()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "LocalResourceCount");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int32 ResourceCount()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ResourceCount");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="assignmentToDrag">object assignmentToDrag</param>
		/// <param name="destinationResource">optional object destinationResource</param>
		/// <param name="destinationTime">optional object destinationTime</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSProject", 11,14)]
		public virtual void RSVDragSimulator(object assignmentToDrag, object destinationResource, object destinationTime)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RSVDragSimulator", assignmentToDrag, destinationResource, destinationTime);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="assignmentToDrag">object assignmentToDrag</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual void RSVDragSimulator(object assignmentToDrag)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RSVDragSimulator", assignmentToDrag);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="assignmentToDrag">object assignmentToDrag</param>
		/// <param name="destinationResource">optional object destinationResource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual void RSVDragSimulator(object assignmentToDrag, object destinationResource)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RSVDragSimulator", assignmentToDrag, destinationResource);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="customUIXML">string customUIXML</param>
		[SupportByVersion("MSProject", 11,14)]
		public virtual void SetCustomUI(string customUIXML)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetCustomUI", customUIXML);
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
		public virtual void ExportAsFixedFormat(string filename, object fileType, object includeDocumentProperties, object includeDocumentMarkup, object archiveFormat, object fromDate, object toDate, object fixedFormatExtClassPtr)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ filename, fileType, includeDocumentProperties, includeDocumentMarkup, archiveFormat, fromDate, toDate, fixedFormatExtClassPtr });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual void ExportAsFixedFormat(string filename)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", filename);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileType">optional NetOffice.MSProjectApi.Enums.PjDocExportType FileType = 0</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual void ExportAsFixedFormat(string filename, object fileType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", filename, fileType);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="fileType">optional NetOffice.MSProjectApi.Enums.PjDocExportType FileType = 0</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,14)]
		public virtual void ExportAsFixedFormat(string filename, object fileType, object includeDocumentProperties)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", filename, fileType, includeDocumentProperties);
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
		public virtual void ExportAsFixedFormat(string filename, object fileType, object includeDocumentProperties, object includeDocumentMarkup)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", filename, fileType, includeDocumentProperties, includeDocumentMarkup);
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
		public virtual void ExportAsFixedFormat(string filename, object fileType, object includeDocumentProperties, object includeDocumentMarkup, object archiveFormat)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ filename, fileType, includeDocumentProperties, includeDocumentMarkup, archiveFormat });
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
		public virtual void ExportAsFixedFormat(string filename, object fileType, object includeDocumentProperties, object includeDocumentMarkup, object archiveFormat, object fromDate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ filename, fileType, includeDocumentProperties, includeDocumentMarkup, archiveFormat, fromDate });
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
		public virtual void ExportAsFixedFormat(string filename, object fileType, object includeDocumentProperties, object includeDocumentMarkup, object archiveFormat, object fromDate, object toDate)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ filename, fileType, includeDocumentProperties, includeDocumentMarkup, archiveFormat, fromDate, toDate });
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public virtual Int32 CheckoutProject()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CheckoutProject");
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public virtual Int32 HideCheckoutMsgBar()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "HideCheckoutMsgBar");
		}

		#endregion

		#pragma warning restore
	}
}


