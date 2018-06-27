using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSProjectApi;

namespace NetOffice.MSProjectApi.Behind
{
	/// <summary>
	/// DispatchInterface Task 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920717(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Task : COMObject, NetOffice.MSProjectApi.Task
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
                    _contractType = typeof(NetOffice.MSProjectApi.Task);
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
                    _type = typeof(Task);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Task() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Work
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Work");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object BaselineWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BaselineWork");			}
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
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object WorkVariance
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "WorkVariance");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object RemainingWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "RemainingWork");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "RemainingWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Cost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Cost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public virtual object ActualCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ActualCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ActualCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public virtual object SV
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "SV");
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
		public virtual object ConstraintType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ConstraintType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ConstraintType", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object ConstraintDate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ConstraintDate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ConstraintDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Critical
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Critical");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object LevelingDelay
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "LevelingDelay");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "LevelingDelay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public virtual Int32 ID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ID");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Milestone
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Milestone");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Milestone", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public virtual string Subproject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Subproject");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Subproject", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object BaselineDuration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BaselineDuration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BaselineDuration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object ActualDuration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ActualDuration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ActualDuration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Duration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object DurationVariance
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DurationVariance");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object RemainingDuration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "RemainingDuration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "RemainingDuration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object PercentComplete
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PercentComplete");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "PercentComplete", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public virtual object Start
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Start");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Finish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Finish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public virtual object LateFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "LateFinish");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object ActualStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ActualStart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ActualStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object ActualFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ActualFinish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ActualFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public virtual object FinishVariance
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FinishVariance");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public virtual string Project
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Project");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int16 OutlineLevel
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "OutlineLevel");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OutlineLevel", value);
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public virtual object Created
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Created");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string UniqueIDPredecessors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "UniqueIDPredecessors");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UniqueIDPredecessors", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string UniqueIDSuccessors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "UniqueIDSuccessors");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UniqueIDSuccessors", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public virtual object LinkedFields
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "LinkedFields");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Resume
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Resume");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Resume", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Stop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Stop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Stop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object ResumeNoEarlierThan
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ResumeNoEarlierThan");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ResumeNoEarlierThan", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string OutlineNumber
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OutlineNumber");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public virtual object UpdateNeeded
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "UpdateNeeded");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string ResourceGroup
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ResourceGroup");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object ACWP
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ACWP");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjTaskFixedType Type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjTaskFixedType>(this, "Type");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Recurring
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Recurring");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EffortDriven
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EffortDriven");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EffortDriven", value);
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
		public virtual NetOffice.MSProjectApi.Tasks PredecessorTasks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Tasks>(this, "PredecessorTasks", typeof(NetOffice.MSProjectApi.Tasks));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Tasks SuccessorTasks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Tasks>(this, "SuccessorTasks", typeof(NetOffice.MSProjectApi.Tasks));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object OvertimeWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "OvertimeWork");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object ActualOvertimeWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ActualOvertimeWork");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object RemainingOvertimeWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "RemainingOvertimeWork");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object RegularWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "RegularWork");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object OvertimeCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "OvertimeCost");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object ActualOvertimeCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ActualOvertimeCost");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object RemainingOvertimeCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "RemainingOvertimeCost");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjAccrueAt FixedCostAccrual
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "FixedCostAccrual");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Assignments Assignments
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Assignments>(this, "Assignments", typeof(NetOffice.MSProjectApi.Assignments));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Parent", value);
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Hyperlink
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Hyperlink");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Hyperlink", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string HyperlinkAddress
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HyperlinkAddress");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HyperlinkAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string HyperlinkSubAddress
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HyperlinkSubAddress");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HyperlinkSubAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string HyperlinkHREF
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HyperlinkHREF");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HyperlinkHREF", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Overallocated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Overallocated");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.SplitParts SplitParts
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.SplitParts>(this, "SplitParts", typeof(NetOffice.MSProjectApi.SplitParts));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object ExternalTask
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ExternalTask");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Task OutlineParent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Task>(this, "OutlineParent", typeof(NetOffice.MSProjectApi.Task));
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object SubProjectReadOnly
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "SubProjectReadOnly");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "SubProjectReadOnly", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object ResponsePending
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ResponsePending");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object TeamStatusPending
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "TeamStatusPending");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object LevelingCanSplit
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "LevelingCanSplit");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "LevelingCanSplit", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object LevelIndividualAssignments
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "LevelIndividualAssignments");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "LevelIndividualAssignments", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Cost4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Cost4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Cost4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Cost5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Cost5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Cost5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Cost6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Cost6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Cost6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Cost7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Cost7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Cost7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Cost8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Cost8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Cost8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Cost9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Cost9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Cost9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Cost10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Cost10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Cost10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Date1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Date1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Date1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Date2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Date2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Date2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Date3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Date3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Date3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Date4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Date4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Date4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Date5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Date5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Date5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Date6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Date6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Date6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Date7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Date7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Date7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Date8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Date8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Date8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Date9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Date9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Date9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Date10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Date10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Date10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Duration4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Duration5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Duration6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Duration7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Duration8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Duration9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Duration10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Start6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Start6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Start6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Finish6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Finish6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Finish6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Start7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Start7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Start7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Finish7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Finish7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Finish7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Start8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Start8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Start8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Finish8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Finish8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Finish8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Start9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Start9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Start9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Finish9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Finish9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Finish9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Start10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Start10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Start10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Finish10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Finish10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Finish10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Flag11
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag11");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Flag12
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag12");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Flag13
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag13");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Flag14
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag14");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Flag15
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag15");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Flag16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag16");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Flag17
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag17");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Flag18
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag18");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Flag19
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag19");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Flag20
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Flag20");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Flag20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double Number6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double Number7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double Number8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double Number9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double Number10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double Number11
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number11");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double Number12
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number12");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double Number13
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number13");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double Number14
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number14");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double Number15
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number15");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double Number16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number16");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double Number17
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number17");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double Number18
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number18");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double Number19
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number19");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double Number20
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Number20");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Number20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text11
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text11");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text12
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text12");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text13
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text13");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text14
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text14");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text15
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text15");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text16");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text17
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text17");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text18
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text18");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text19
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text19");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text20
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text20");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text21
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text21");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text22
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text22");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text23
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text23");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text24
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text24");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text25
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text25");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text26
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text26");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text27
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text27");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text28
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text28");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text29
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text29");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Text30
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text30");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text30", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string ResourcePhonetics
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ResourcePhonetics");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int32 Index
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object PreleveledStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PreleveledStart");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object PreleveledFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PreleveledFinish");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Predecessors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Predecessors");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Predecessors", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Successors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Successors");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Successors", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string ResourceNames
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ResourceNames");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ResourceNames", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string ResourceInitials
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ResourceInitials");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ResourceInitials", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Estimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Estimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object IgnoreResourceCalendar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "IgnoreResourceCalendar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "IgnoreResourceCalendar", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Calendar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Calendar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Calendar", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Duration1Estimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration1Estimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration1Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Duration2Estimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration2Estimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration2Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Duration3Estimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration3Estimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration3Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Duration4Estimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration4Estimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration4Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Duration5Estimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration5Estimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration5Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Duration6Estimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration6Estimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration6Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Duration7Estimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration7Estimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration7Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Duration8Estimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration8Estimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration8Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Duration9Estimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration9Estimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration9Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Duration10Estimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Duration10Estimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Duration10Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object BaselineDurationEstimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BaselineDurationEstimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BaselineDurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string OutlineCode1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OutlineCode1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OutlineCode1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string OutlineCode2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OutlineCode2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OutlineCode2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string OutlineCode3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OutlineCode3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OutlineCode3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string OutlineCode4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OutlineCode4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OutlineCode4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string OutlineCode5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OutlineCode5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OutlineCode5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string OutlineCode6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OutlineCode6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OutlineCode6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string OutlineCode7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OutlineCode7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OutlineCode7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string OutlineCode8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OutlineCode8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OutlineCode8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string OutlineCode9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OutlineCode9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OutlineCode9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string OutlineCode10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OutlineCode10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OutlineCode10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Deadline
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Deadline");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Deadline", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object StartSlack
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "StartSlack");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object FinishSlack
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FinishSlack");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object VAC
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "VAC");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.TaskDependencies TaskDependencies
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.TaskDependencies>(this, "TaskDependencies", typeof(NetOffice.MSProjectApi.TaskDependencies));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object GroupBySummary
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "GroupBySummary");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string WBSPredecessors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "WBSPredecessors");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string WBSSuccessors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "WBSSuccessors");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string HyperlinkScreenTip
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HyperlinkScreenTip");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HyperlinkScreenTip", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline1Start
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline1Start");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline1Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline1Finish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline1Finish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline1Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline1Cost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline1Cost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline1Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline1Work
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline1Work");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline1Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline1Duration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline1Duration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline1Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline2Start
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline2Start");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline2Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline2Finish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline2Finish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline2Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline2Cost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline2Cost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline2Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline2Work
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline2Work");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline2Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline2Duration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline2Duration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline2Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline3Start
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline3Start");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline3Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline3Finish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline3Finish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline3Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline3Cost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline3Cost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline3Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline3Work
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline3Work");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline3Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline3Duration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline3Duration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline3Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline4Start
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline4Start");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline4Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline4Finish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline4Finish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline4Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline4Cost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline4Cost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline4Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline4Work
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline4Work");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline4Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline4Duration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline4Duration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline4Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline5Start
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline5Start");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline5Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline5Finish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline5Finish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline5Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline5Cost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline5Cost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline5Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline5Work
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline5Work");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline5Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline5Duration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline5Duration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline5Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double CPI
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "CPI");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double SPI
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "SPI");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object CVPercent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "CVPercent");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object SVPercent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "SVPercent");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EAC
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EAC");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double TCPI
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "TCPI");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjStatusType Status
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjStatusType>(this, "Status");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline6Start
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline6Start");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline6Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline6Finish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline6Finish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline6Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline6Cost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline6Cost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline6Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline6Work
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline6Work");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline6Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline6Duration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline6Duration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline6Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline7Start
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline7Start");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline7Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline7Finish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline7Finish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline7Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline7Cost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline7Cost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline7Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline7Work
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline7Work");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline7Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline7Duration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline7Duration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline7Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline8Start
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline8Start");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline8Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline8Finish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline8Finish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline8Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline8Cost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline8Cost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline8Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline8Work
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline8Work");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline8Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline8Duration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline8Duration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline8Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline9Start
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline9Start");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline9Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline9Finish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline9Finish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline9Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline9Cost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline9Cost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline9Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline9Work
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline9Work");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline9Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline9Duration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline9Duration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline9Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline10Start
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline10Start");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline10Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline10Finish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline10Finish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline10Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline10Cost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline10Cost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline10Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline10Work
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline10Work");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline10Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline10Duration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline10Duration");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline10Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseCost1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseCost1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseCost1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseCost2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseCost2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseCost2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseCost3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseCost3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseCost3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseCost4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseCost4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseCost4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseCost5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseCost5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseCost5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseCost6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseCost6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseCost6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseCost7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseCost7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseCost7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseCost8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseCost8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseCost8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseCost9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseCost9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseCost9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseCost10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseCost10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseCost10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate11
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate11");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate12
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate12");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate13
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate13");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate14
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate14");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate15
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate15");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate16");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate17
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate17");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate18
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate18");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate19
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate19");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate20
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate20");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate21
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate21");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate22
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate22");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate23
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate23");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate24
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate24");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate25
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate25");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate26
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate26");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate27
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate27");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate28
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate28");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate29
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate29");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDate30
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDate30");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDate30", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDuration1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDuration1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDuration1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDuration2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDuration2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDuration2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDuration3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDuration3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDuration3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDuration4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDuration4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDuration4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDuration5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDuration5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDuration5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDuration6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDuration6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDuration6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDuration7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDuration7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDuration7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDuration8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDuration8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDuration8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDuration9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDuration9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDuration9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseDuration10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseDuration10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseDuration10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag11
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag11");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag12
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag12");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag13
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag13");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag14
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag14");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag15
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag15");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag16");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag17
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag17");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag18
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag18");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag19
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag19");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseFlag20
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseFlag20");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseFlag20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber11
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber11");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber12
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber12");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber13
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber13");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber14
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber14");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber15
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber15");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber16");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber17
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber17");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber18
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber18");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber19
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber19");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber20
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber20");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber21
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber21");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber22
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber22");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber23
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber23");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber24
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber24");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber25
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber25");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber26
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber26");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber27
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber27");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber28
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber28");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber29
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber29");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber30
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber30");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber30", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber31
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber31");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber31", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber32
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber32");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber32", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber33
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber33");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber33", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber34
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber34");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber34", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber35
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber35");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber35", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber36
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber36");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber36", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber37
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber37");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber37", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber38
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber38");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber38", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber39
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber39");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber39", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseNumber40
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseNumber40");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseNumber40", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode11
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode11");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode12
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode12");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode13
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode13");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode14
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode14");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode15
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode15");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode16");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode17
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode17");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode18
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode18");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode19
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode19");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode20
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode20");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode21
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode21");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode22
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode22");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode23
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode23");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode24
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode24");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode25
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode25");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode26
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode26");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode27
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode27");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode28
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode28");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode29
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode29");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseOutlineCode30
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode30");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseOutlineCode30", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText11
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText11");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText12
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText12");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText13
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText13");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText14
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText14");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText15
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText15");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText16");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText17
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText17");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText18
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText18");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText19
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText19");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText20
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText20");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText21
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText21");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText22
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText22");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText23
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText23");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText24
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText24");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText25
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText25");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText26
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText26");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText27
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText27");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText28
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText28");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText29
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText29");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText30
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText30");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText30", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText31
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText31");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText31", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText32
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText32");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText32", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText33
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText33");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText33", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText34
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText34");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText34", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText35
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText35");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText35", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText36
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText36");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText36", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText37
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText37");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText37", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText38
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText38");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText38", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText39
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText39");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText39", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseText40
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseText40");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseText40", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline1DurationEstimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline1DurationEstimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline1DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline2DurationEstimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline2DurationEstimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline2DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline3DurationEstimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline3DurationEstimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline3DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline4DurationEstimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline4DurationEstimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline4DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline5DurationEstimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline5DurationEstimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline5DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline6DurationEstimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline6DurationEstimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline6DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline7DurationEstimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline7DurationEstimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline7DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline8DurationEstimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline8DurationEstimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline8DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline9DurationEstimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline9DurationEstimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline9DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline10DurationEstimated
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline10DurationEstimated");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline10DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object PhysicalPercentComplete
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PhysicalPercentComplete");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "PhysicalPercentComplete", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjEarnedValueMethod EarnedValueMethod
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjEarnedValueMethod>(this, "EarnedValueMethod");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "EarnedValueMethod", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectCost1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectCost1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectCost2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectCost2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectCost3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectCost3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectCost4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectCost4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectCost5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectCost5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectCost6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectCost6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectCost7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectCost7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectCost8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectCost8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectCost9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectCost9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectCost10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectCost10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate11
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate11");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate12
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate12");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate13
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate13");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate14
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate14");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate15
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate15");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate16");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate17
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate17");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate18
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate18");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate19
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate19");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate20
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate20");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate21
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate21");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate22
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate22");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate23
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate23");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate24
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate24");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate25
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate25");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate26
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate26");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate27
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate27");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate28
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate28");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate29
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate29");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDate30
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate30");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDate30", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDuration1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDuration2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDuration3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDuration4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDuration5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDuration6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDuration7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDuration8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDuration9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectDuration10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode11
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode11");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode12
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode12");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode13
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode13");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode14
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode14");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode15
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode15");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode16");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode17
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode17");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode18
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode18");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode19
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode19");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode20
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode20");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode21
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode21");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode22
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode22");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode23
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode23");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode24
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode24");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode25
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode25");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode26
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode26");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode27
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode27");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode28
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode28");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode29
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode29");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectOutlineCode30
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode30");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode30", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag11
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag11");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag12
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag12");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag13
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag13");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag14
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag14");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag15
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag15");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag16");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag17
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag17");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag18
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag18");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag19
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag19");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object EnterpriseProjectFlag20
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag20");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber11
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber11");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber12
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber12");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber13
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber13");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber14
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber14");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber15
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber15");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber16");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber17
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber17");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber18
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber18");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber19
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber19");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber20
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber20");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber21
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber21");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber22
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber22");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber23
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber23");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber24
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber24");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber25
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber25");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber26
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber26");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber27
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber27");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber28
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber28");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber29
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber29");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber30
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber30");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber30", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber31
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber31");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber31", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber32
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber32");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber32", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber33
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber33");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber33", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber34
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber34");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber34", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber35
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber35");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber35", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber36
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber36");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber36", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber37
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber37");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber37", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber38
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber38");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber38", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber39
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber39");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber39", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Double EnterpriseProjectNumber40
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber40");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectNumber40", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText4
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText4");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText5
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText5");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText6
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText6");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText7
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText7");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText8
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText8");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText9
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText9");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText10
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText10");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText11
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText11");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText12
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText12");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText13
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText13");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText14
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText14");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText15
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText15");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText16");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText17
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText17");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText18
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText18");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText19
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText19");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText20
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText20");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText21
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText21");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText22
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText22");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText23
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText23");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText24
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText24");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText25
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText25");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText26
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText26");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText27
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText27");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText28
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText28");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText29
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText29");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText30
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText30");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText30", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText31
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText31");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText31", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText32
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText32");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText32", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText33
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText33");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText33", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText34
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText34");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText34", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText35
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText35");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText35", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText36
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText36");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText36", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText37
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText37");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText37", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText38
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText38");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText38", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText39
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText39");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText39", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string EnterpriseProjectText40
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EnterpriseProjectText40");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnterpriseProjectText40", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object ActualWorkProtected
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ActualWorkProtected");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ActualWorkProtected", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object ActualOvertimeWorkProtected
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ActualOvertimeWorkProtected");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ActualOvertimeWorkProtected", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object BaselineFixedCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BaselineFixedCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BaselineFixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline1FixedCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline1FixedCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline1FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline2FixedCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline2FixedCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline2FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline3FixedCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline3FixedCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline3FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline4FixedCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline4FixedCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline4FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline5FixedCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline5FixedCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline5FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline6FixedCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline6FixedCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline6FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline7FixedCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline7FixedCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline7FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline8FixedCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline8FixedCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline8FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline9FixedCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline9FixedCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline9FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline10FixedCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline10FixedCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline10FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string Guid
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Guid");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string CalendarGuid
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CalendarGuid");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string DeliverableGuid
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DeliverableGuid");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DeliverableGuid", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int16 DeliverableType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "DeliverableType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DeliverableType", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object DeliverableStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DeliverableStart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object DeliverableFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DeliverableFinish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object IsPublished
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "IsPublished");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "IsPublished", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string StatusManagerName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "StatusManagerName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StatusManagerName", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string ErrorMessage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ErrorMessage");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object BudgetWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BudgetWork");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object BudgetCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BudgetCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjAccrueAt BaselineFixedCostAccrual
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "BaselineFixedCostAccrual");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "BaselineFixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object BaselineDeliverableStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BaselineDeliverableStart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BaselineDeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object BaselineDeliverableFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BaselineDeliverableFinish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BaselineDeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object BaselineBudgetWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BaselineBudgetWork");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BaselineBudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object BaselineBudgetCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BaselineBudgetCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BaselineBudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline1FixedCostAccrual
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline1FixedCostAccrual");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Baseline1FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline1DeliverableStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline1DeliverableStart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline1DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline1DeliverableFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline1DeliverableFinish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline1DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline1BudgetWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline1BudgetWork");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline1BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline1BudgetCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline1BudgetCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline1BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline2FixedCostAccrual
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline2FixedCostAccrual");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Baseline2FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline2DeliverableStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline2DeliverableStart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline2DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline2DeliverableFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline2DeliverableFinish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline2DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline2BudgetWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline2BudgetWork");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline2BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline2BudgetCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline2BudgetCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline2BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline3FixedCostAccrual
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline3FixedCostAccrual");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Baseline3FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline3DeliverableStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline3DeliverableStart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline3DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline3DeliverableFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline3DeliverableFinish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline3DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline3BudgetWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline3BudgetWork");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline3BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline3BudgetCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline3BudgetCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline3BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline4FixedCostAccrual
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline4FixedCostAccrual");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Baseline4FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline4DeliverableStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline4DeliverableStart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline4DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline4DeliverableFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline4DeliverableFinish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline4DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline4BudgetWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline4BudgetWork");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline4BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline4BudgetCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline4BudgetCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline4BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline5FixedCostAccrual
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline5FixedCostAccrual");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Baseline5FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline5DeliverableStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline5DeliverableStart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline5DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline5DeliverableFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline5DeliverableFinish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline5DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline5BudgetWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline5BudgetWork");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline5BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline5BudgetCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline5BudgetCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline5BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline6FixedCostAccrual
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline6FixedCostAccrual");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Baseline6FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline6DeliverableStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline6DeliverableStart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline6DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline6DeliverableFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline6DeliverableFinish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline6DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline6BudgetWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline6BudgetWork");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline6BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline6BudgetCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline6BudgetCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline6BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline7FixedCostAccrual
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline7FixedCostAccrual");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Baseline7FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline7DeliverableStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline7DeliverableStart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline7DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline7DeliverableFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline7DeliverableFinish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline7DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline7BudgetWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline7BudgetWork");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline7BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline7BudgetCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline7BudgetCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline7BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline8FixedCostAccrual
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline8FixedCostAccrual");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Baseline8FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline8DeliverableStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline8DeliverableStart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline8DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline8DeliverableFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline8DeliverableFinish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline8DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline8BudgetWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline8BudgetWork");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline8BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline8BudgetCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline8BudgetCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline8BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline9FixedCostAccrual
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline9FixedCostAccrual");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Baseline9FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline9DeliverableStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline9DeliverableStart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline9DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline9DeliverableFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline9DeliverableFinish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline9DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline9BudgetWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline9BudgetWork");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline9BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline9BudgetCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline9BudgetCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline9BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline10FixedCostAccrual
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline10FixedCostAccrual");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Baseline10FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline10DeliverableStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline10DeliverableStart");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline10DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline10DeliverableFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline10DeliverableFinish");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline10DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline10BudgetWork
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline10BudgetWork");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline10BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual object Baseline10BudgetCost
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Baseline10BudgetCost");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Baseline10BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual Int32 RecalcFlags
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RecalcFlags");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.StartDriver StartDriver
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.StartDriver>(this, "StartDriver", typeof(NetOffice.MSProjectApi.StartDriver));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string DeliverableName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DeliverableName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DeliverableName", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual object Active
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Active");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Active", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual object Manual
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Manual");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Manual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual object Placeholder
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Placeholder");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual object Warning
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Warning");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string StartText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "StartText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string FinishText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FinishText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string DurationText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DurationText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864567(v=office.14).aspx </remarks>
		[SupportByVersion("MSProject", 11,14)]
		public virtual object IsStartValid
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "IsStartValid");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861711(v=office.14).aspx </remarks>
		[SupportByVersion("MSProject", 11,14)]
		public virtual object IsFinishValid
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "IsFinishValid");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865936(v=office.14).aspx </remarks>
		[SupportByVersion("MSProject", 11,14)]
		public virtual object IsDurationValid
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "IsDurationValid");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string BaselineStartText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BaselineStartText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BaselineStartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string BaselineFinishText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BaselineFinishText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BaselineFinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string BaselineDurationText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BaselineDurationText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BaselineDurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline1StartText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline1StartText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline1StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline1FinishText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline1FinishText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline1FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline1DurationText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline1DurationText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline1DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline2StartText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline2StartText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline2StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline2FinishText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline2FinishText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline2FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline2DurationText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline2DurationText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline2DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline3StartText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline3StartText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline3StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline3FinishText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline3FinishText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline3FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline3DurationText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline3DurationText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline3DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline4StartText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline4StartText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline4StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline4FinishText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline4FinishText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline4FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline4DurationText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline4DurationText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline4DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline5StartText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline5StartText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline5StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline5FinishText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline5FinishText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline5FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline5DurationText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline5DurationText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline5DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline6StartText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline6StartText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline6StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline6FinishText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline6FinishText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline6FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline6DurationText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline6DurationText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline6DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline7StartText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline7StartText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline7StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline7FinishText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline7FinishText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline7FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline7DurationText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline7DurationText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline7DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline8StartText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline8StartText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline8StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline8FinishText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline8FinishText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline8FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline8DurationText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline8DurationText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline8DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline9StartText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline9StartText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline9StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline9FinishText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline9FinishText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline9FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline9DurationText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline9DurationText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline9DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline10StartText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline10StartText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline10StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline10FinishText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline10FinishText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline10FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual string Baseline10DurationText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Baseline10DurationText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Baseline10DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual object IgnoreWarnings
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "IgnoreWarnings");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "IgnoreWarnings", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual NetOffice.MSProjectApi.Calendar CalendarObject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Calendar>(this, "CalendarObject", typeof(NetOffice.MSProjectApi.Calendar));
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual object ScheduledStart
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ScheduledStart");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual object ScheduledFinish
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ScheduledFinish");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public virtual object ScheduledDuration
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ScheduledDuration");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public virtual object PathDrivingPredecessor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PathDrivingPredecessor");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public virtual object PathPredecessor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PathPredecessor");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public virtual object PathDrivenSuccessor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PathDrivenSuccessor");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public virtual object PathSuccessor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PathSuccessor");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="fieldID">NetOffice.MSProjectApi.Enums.PjField fieldID</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual string GetField(NetOffice.MSProjectApi.Enums.PjField fieldID)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetField", fieldID);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="fieldID">NetOffice.MSProjectApi.Enums.PjField fieldID</param>
		/// <param name="value">string value</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void SetField(NetOffice.MSProjectApi.Enums.PjField fieldID, string value)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetField", fieldID, value);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void Delete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
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
		/// <param name="startDate">object startDate</param>
		/// <param name="endDate">object endDate</param>
		/// <param name="type">optional NetOffice.MSProjectApi.Enums.PjTaskTimescaledData Type = 0</param>
		/// <param name="timeScaleUnit">optional NetOffice.MSProjectApi.Enums.PjTimescaleUnit TimeScaleUnit = 3</param>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.TimeScaleValues TimeScaleData(object startDate, object endDate, object type, object timeScaleUnit, object count)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TimeScaleValues>(this, "TimeScaleData", typeof(NetOffice.MSProjectApi.TimeScaleValues), new object[]{ startDate, endDate, type, timeScaleUnit, count });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="startDate">object startDate</param>
		/// <param name="endDate">object endDate</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.TimeScaleValues TimeScaleData(object startDate, object endDate)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TimeScaleValues>(this, "TimeScaleData", typeof(NetOffice.MSProjectApi.TimeScaleValues), startDate, endDate);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="startDate">object startDate</param>
		/// <param name="endDate">object endDate</param>
		/// <param name="type">optional NetOffice.MSProjectApi.Enums.PjTaskTimescaledData Type = 0</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.TimeScaleValues TimeScaleData(object startDate, object endDate, object type)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TimeScaleValues>(this, "TimeScaleData", typeof(NetOffice.MSProjectApi.TimeScaleValues), startDate, endDate, type);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="startDate">object startDate</param>
		/// <param name="endDate">object endDate</param>
		/// <param name="type">optional NetOffice.MSProjectApi.Enums.PjTaskTimescaledData Type = 0</param>
		/// <param name="timeScaleUnit">optional NetOffice.MSProjectApi.Enums.PjTimescaleUnit TimeScaleUnit = 3</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.TimeScaleValues TimeScaleData(object startDate, object endDate, object type, object timeScaleUnit)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TimeScaleValues>(this, "TimeScaleData", typeof(NetOffice.MSProjectApi.TimeScaleValues), startDate, endDate, type, timeScaleUnit);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tasks">object tasks</param>
		/// <param name="link">optional NetOffice.MSProjectApi.Enums.PjTaskLinkType Link = 1</param>
		/// <param name="lag">optional object lag</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void LinkPredecessors(object tasks, object link, object lag)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LinkPredecessors", tasks, link, lag);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tasks">object tasks</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void LinkPredecessors(object tasks)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LinkPredecessors", tasks);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tasks">object tasks</param>
		/// <param name="link">optional NetOffice.MSProjectApi.Enums.PjTaskLinkType Link = 1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void LinkPredecessors(object tasks, object link)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LinkPredecessors", tasks, link);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tasks">object tasks</param>
		/// <param name="link">optional NetOffice.MSProjectApi.Enums.PjTaskLinkType Link = 1</param>
		/// <param name="lag">optional object lag</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void LinkSuccessors(object tasks, object link, object lag)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LinkSuccessors", tasks, link, lag);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tasks">object tasks</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void LinkSuccessors(object tasks)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LinkSuccessors", tasks);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tasks">object tasks</param>
		/// <param name="link">optional NetOffice.MSProjectApi.Enums.PjTaskLinkType Link = 1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void LinkSuccessors(object tasks, object link)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LinkSuccessors", tasks, link);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tasks">object tasks</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void UnlinkPredecessors(object tasks)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UnlinkPredecessors", tasks);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tasks">object tasks</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void UnlinkSuccessors(object tasks)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UnlinkSuccessors", tasks);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void OutlineIndent()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutlineIndent");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void OutlineOutdent()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutlineOutdent");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void OutlineHideSubTasks()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutlineHideSubTasks");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void OutlineShowSubTasks()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutlineShowSubTasks");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void OutlineShowAllTasks()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutlineShowAllTasks");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="startSplitOn">object startSplitOn</param>
		/// <param name="endSplitOn">object endSplitOn</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual void Split(object startSplitOn, object endSplitOn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Split", startSplitOn, endSplitOn);
		}

		#endregion

		#pragma warning restore
	}
}


