using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface Task 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920717(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Task : COMObject
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
                    _type = typeof(Task);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Task(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Task(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Task(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Task(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Task(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Task(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Task() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Task(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Work
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Work");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object WorkVariance
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "WorkVariance");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object RemainingWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "RemainingWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "RemainingWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Cost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Cost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public object ActualCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ActualCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ActualCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public object SV
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "SV");
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
		public object ConstraintType
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ConstraintType");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ConstraintType", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object ConstraintDate
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ConstraintDate");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ConstraintDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Critical
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Critical");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object LevelingDelay
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LevelingDelay");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "LevelingDelay", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public Int32 ID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ID");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Milestone
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Milestone");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Milestone", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public string Subproject
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Subproject");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Subproject", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object BaselineDuration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BaselineDuration");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BaselineDuration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object ActualDuration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ActualDuration");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ActualDuration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Duration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object DurationVariance
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DurationVariance");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object RemainingDuration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "RemainingDuration");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "RemainingDuration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object PercentComplete
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PercentComplete");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PercentComplete", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public object Start
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Start");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Finish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Finish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public object LateFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LateFinish");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object ActualStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ActualStart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ActualStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object ActualFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ActualFinish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ActualFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public object FinishVariance
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FinishVariance");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public string Project
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Project");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int16 OutlineLevel
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "OutlineLevel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OutlineLevel", value);
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public object Created
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Created");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string UniqueIDPredecessors
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UniqueIDPredecessors");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UniqueIDPredecessors", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string UniqueIDSuccessors
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UniqueIDSuccessors");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UniqueIDSuccessors", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public object LinkedFields
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LinkedFields");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Resume
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Resume");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Resume", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Stop
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Stop");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Stop", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ResumeNoEarlierThan
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ResumeNoEarlierThan");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ResumeNoEarlierThan", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string OutlineNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OutlineNumber");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		public object UpdateNeeded
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "UpdateNeeded");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
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
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string ResourceGroup
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ResourceGroup");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object ACWP
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ACWP");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjTaskFixedType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjTaskFixedType>(this, "Type");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Recurring
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Recurring");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EffortDriven
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EffortDriven");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EffortDriven", value);
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
		public NetOffice.MSProjectApi.Tasks PredecessorTasks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Tasks>(this, "PredecessorTasks", NetOffice.MSProjectApi.Tasks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Tasks SuccessorTasks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Tasks>(this, "SuccessorTasks", NetOffice.MSProjectApi.Tasks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object OvertimeWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "OvertimeWork");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object ActualOvertimeWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ActualOvertimeWork");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object RemainingOvertimeWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "RemainingOvertimeWork");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object RegularWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "RegularWork");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object OvertimeCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "OvertimeCost");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object ActualOvertimeCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ActualOvertimeCost");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object RemainingOvertimeCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "RemainingOvertimeCost");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjAccrueAt FixedCostAccrual
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "FixedCostAccrual");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Assignments Assignments
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Assignments>(this, "Assignments", NetOffice.MSProjectApi.Assignments.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Parent", value);
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Hyperlink
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Hyperlink");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Hyperlink", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string HyperlinkAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HyperlinkAddress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HyperlinkAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string HyperlinkSubAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HyperlinkSubAddress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HyperlinkSubAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string HyperlinkHREF
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HyperlinkHREF");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HyperlinkHREF", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Overallocated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Overallocated");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.SplitParts SplitParts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.SplitParts>(this, "SplitParts", NetOffice.MSProjectApi.SplitParts.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object ExternalTask
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ExternalTask");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Task OutlineParent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Task>(this, "OutlineParent", NetOffice.MSProjectApi.Task.LateBindingApiWrapperType);
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object SubProjectReadOnly
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "SubProjectReadOnly");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "SubProjectReadOnly", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object ResponsePending
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ResponsePending");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object TeamStatusPending
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TeamStatusPending");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object LevelingCanSplit
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LevelingCanSplit");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "LevelingCanSplit", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object LevelIndividualAssignments
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LevelIndividualAssignments");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "LevelIndividualAssignments", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Cost4
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Cost4");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Cost4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Cost5
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Cost5");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Cost5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Cost6
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Cost6");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Cost6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Cost7
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Cost7");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Cost7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Cost8
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Cost8");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Cost8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Cost9
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Cost9");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Cost9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Cost10
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Cost10");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Cost10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Date1
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Date1");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Date1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Date2
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Date2");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Date2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Date3
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Date3");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Date3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Date4
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Date4");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Date4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Date5
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Date5");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Date5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Date6
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Date6");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Date6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Date7
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Date7");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Date7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Date8
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Date8");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Date8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Date9
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Date9");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Date9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Date10
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Date10");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Date10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Duration4
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration4");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Duration5
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration5");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Duration6
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration6");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Duration7
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration7");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Duration8
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration8");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Duration9
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration9");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Duration10
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration10");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Start6
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Start6");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Start6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Finish6
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Finish6");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Finish6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Start7
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Start7");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Start7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Finish7
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Finish7");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Finish7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Start8
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Start8");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Start8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Finish8
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Finish8");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Finish8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Start9
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Start9");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Start9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Finish9
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Finish9");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Finish9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Start10
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Start10");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Start10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Finish10
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Finish10");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Finish10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Flag11
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag11");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Flag12
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag12");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Flag13
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag13");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Flag14
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag14");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Flag15
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag15");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Flag16
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag16");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Flag17
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag17");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Flag18
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag18");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Flag19
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag19");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Flag20
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Flag20");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Flag20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double Number6
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number6");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double Number7
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number7");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double Number8
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number8");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double Number9
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number9");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double Number10
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number10");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double Number11
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number11");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double Number12
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number12");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double Number13
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number13");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double Number14
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number14");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double Number15
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number15");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double Number16
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number16");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double Number17
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number17");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double Number18
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number18");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double Number19
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number19");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double Number20
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Number20");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Number20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text11
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text11");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text12
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text12");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text13
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text13");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text14
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text14");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text15
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text15");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text16
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text16");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text17
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text17");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text18
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text18");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text19
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text19");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text20
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text20");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text21
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text21");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text22
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text22");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text23
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text23");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text24
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text24");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text25
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text25");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text26
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text26");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text27
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text27");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text28
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text28");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text29
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text29");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Text30
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text30");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text30", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string ResourcePhonetics
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ResourcePhonetics");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 Index
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object PreleveledStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PreleveledStart");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object PreleveledFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PreleveledFinish");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Predecessors
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Predecessors");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Predecessors", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Successors
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Successors");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Successors", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string ResourceNames
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ResourceNames");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ResourceNames", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string ResourceInitials
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ResourceInitials");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ResourceInitials", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Estimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Estimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object IgnoreResourceCalendar
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "IgnoreResourceCalendar");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "IgnoreResourceCalendar", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Calendar
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Calendar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Calendar", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Duration1Estimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration1Estimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration1Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Duration2Estimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration2Estimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration2Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Duration3Estimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration3Estimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration3Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Duration4Estimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration4Estimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration4Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Duration5Estimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration5Estimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration5Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Duration6Estimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration6Estimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration6Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Duration7Estimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration7Estimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration7Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Duration8Estimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration8Estimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration8Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Duration9Estimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration9Estimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration9Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Duration10Estimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Duration10Estimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Duration10Estimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object BaselineDurationEstimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BaselineDurationEstimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BaselineDurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string OutlineCode1
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OutlineCode1");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OutlineCode1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string OutlineCode2
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OutlineCode2");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OutlineCode2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string OutlineCode3
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OutlineCode3");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OutlineCode3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string OutlineCode4
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OutlineCode4");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OutlineCode4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string OutlineCode5
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OutlineCode5");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OutlineCode5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string OutlineCode6
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OutlineCode6");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OutlineCode6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string OutlineCode7
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OutlineCode7");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OutlineCode7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string OutlineCode8
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OutlineCode8");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OutlineCode8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string OutlineCode9
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OutlineCode9");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OutlineCode9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string OutlineCode10
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OutlineCode10");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OutlineCode10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Deadline
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Deadline");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Deadline", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object StartSlack
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "StartSlack");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object FinishSlack
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FinishSlack");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object VAC
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "VAC");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TaskDependencies TaskDependencies
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.TaskDependencies>(this, "TaskDependencies", NetOffice.MSProjectApi.TaskDependencies.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object GroupBySummary
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "GroupBySummary");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string WBSPredecessors
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "WBSPredecessors");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string WBSSuccessors
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "WBSSuccessors");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string HyperlinkScreenTip
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HyperlinkScreenTip");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HyperlinkScreenTip", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline1Start
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline1Start");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline1Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline1Finish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline1Finish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline1Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline1Cost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline1Cost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline1Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline1Work
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline1Work");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline1Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline1Duration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline1Duration");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline1Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline2Start
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline2Start");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline2Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline2Finish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline2Finish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline2Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline2Cost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline2Cost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline2Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline2Work
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline2Work");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline2Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline2Duration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline2Duration");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline2Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline3Start
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline3Start");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline3Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline3Finish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline3Finish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline3Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline3Cost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline3Cost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline3Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline3Work
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline3Work");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline3Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline3Duration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline3Duration");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline3Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline4Start
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline4Start");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline4Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline4Finish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline4Finish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline4Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline4Cost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline4Cost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline4Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline4Work
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline4Work");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline4Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline4Duration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline4Duration");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline4Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline5Start
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline5Start");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline5Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline5Finish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline5Finish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline5Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline5Cost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline5Cost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline5Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline5Work
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline5Work");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline5Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline5Duration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline5Duration");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline5Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double CPI
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "CPI");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double SPI
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "SPI");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object CVPercent
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CVPercent");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object SVPercent
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "SVPercent");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EAC
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EAC");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double TCPI
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "TCPI");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjStatusType Status
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjStatusType>(this, "Status");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline6Start
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline6Start");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline6Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline6Finish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline6Finish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline6Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline6Cost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline6Cost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline6Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline6Work
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline6Work");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline6Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline6Duration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline6Duration");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline6Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline7Start
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline7Start");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline7Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline7Finish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline7Finish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline7Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline7Cost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline7Cost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline7Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline7Work
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline7Work");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline7Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline7Duration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline7Duration");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline7Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline8Start
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline8Start");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline8Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline8Finish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline8Finish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline8Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline8Cost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline8Cost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline8Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline8Work
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline8Work");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline8Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline8Duration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline8Duration");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline8Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline9Start
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline9Start");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline9Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline9Finish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline9Finish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline9Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline9Cost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline9Cost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline9Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline9Work
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline9Work");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline9Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline9Duration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline9Duration");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline9Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline10Start
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline10Start");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline10Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline10Finish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline10Finish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline10Finish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline10Cost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline10Cost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline10Cost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline10Work
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline10Work");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline10Work", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline10Duration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline10Duration");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline10Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseCost1
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseCost1");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseCost1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseCost2
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseCost2");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseCost2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseCost3
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseCost3");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseCost3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseCost4
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseCost4");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseCost4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseCost5
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseCost5");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseCost5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseCost6
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseCost6");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseCost6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseCost7
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseCost7");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseCost7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseCost8
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseCost8");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseCost8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseCost9
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseCost9");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseCost9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseCost10
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseCost10");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseCost10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate1
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate1");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate2
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate2");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate3
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate3");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate4
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate4");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate5
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate5");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate6
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate6");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate7
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate7");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate8
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate8");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate9
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate9");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate10
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate10");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate11
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate11");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate12
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate12");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate13
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate13");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate14
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate14");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate15
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate15");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate16
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate16");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate17
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate17");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate18
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate18");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate19
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate19");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate20
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate20");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate21
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate21");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate22
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate22");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate23
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate23");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate24
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate24");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate25
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate25");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate26
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate26");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate27
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate27");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate28
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate28");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate29
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate29");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDate30
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDate30");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDate30", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDuration1
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDuration1");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDuration1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDuration2
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDuration2");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDuration2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDuration3
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDuration3");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDuration3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDuration4
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDuration4");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDuration4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDuration5
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDuration5");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDuration5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDuration6
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDuration6");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDuration6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDuration7
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDuration7");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDuration7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDuration8
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDuration8");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDuration8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDuration9
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDuration9");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDuration9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseDuration10
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseDuration10");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseDuration10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag1
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag1");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag2
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag2");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag3
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag3");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag4
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag4");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag5
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag5");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag6
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag6");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag7
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag7");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag8
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag8");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag9
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag9");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag10
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag10");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag11
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag11");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag12
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag12");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag13
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag13");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag14
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag14");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag15
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag15");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag16
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag16");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag17
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag17");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag18
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag18");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag19
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag19");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseFlag20
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseFlag20");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseFlag20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber1
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber1");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber2
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber2");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber3
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber3");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber4
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber4");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber5
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber5");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber6
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber6");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber7
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber7");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber8
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber8");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber9
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber9");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber10
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber10");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber11
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber11");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber12
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber12");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber13
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber13");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber14
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber14");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber15
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber15");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber16
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber16");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber17
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber17");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber18
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber18");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber19
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber19");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber20
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber20");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber21
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber21");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber22
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber22");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber23
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber23");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber24
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber24");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber25
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber25");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber26
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber26");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber27
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber27");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber28
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber28");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber29
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber29");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber30
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber30");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber30", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber31
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber31");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber31", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber32
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber32");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber32", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber33
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber33");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber33", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber34
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber34");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber34", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber35
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber35");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber35", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber36
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber36");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber36", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber37
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber37");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber37", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber38
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber38");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber38", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber39
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber39");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber39", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseNumber40
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseNumber40");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseNumber40", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode1
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode1");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode2
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode2");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode3
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode3");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode4
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode4");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode5
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode5");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode6
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode6");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode7
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode7");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode8
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode8");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode9
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode9");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode10
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode10");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode11
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode11");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode12
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode12");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode13
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode13");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode14
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode14");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode15
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode15");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode16
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode16");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode17
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode17");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode18
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode18");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode19
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode19");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode20
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode20");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode21
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode21");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode22
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode22");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode23
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode23");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode24
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode24");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode25
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode25");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode26
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode26");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode27
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode27");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode28
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode28");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode29
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode29");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseOutlineCode30
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseOutlineCode30");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseOutlineCode30", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText1
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText1");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText2
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText2");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText3
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText3");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText4
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText4");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText5
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText5");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText6
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText6");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText7
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText7");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText8
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText8");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText9
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText9");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText10
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText10");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText11
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText11");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText12
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText12");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText13
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText13");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText14
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText14");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText15
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText15");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText16
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText16");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText17
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText17");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText18
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText18");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText19
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText19");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText20
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText20");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText21
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText21");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText22
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText22");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText23
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText23");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText24
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText24");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText25
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText25");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText26
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText26");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText27
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText27");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText28
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText28");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText29
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText29");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText30
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText30");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText30", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText31
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText31");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText31", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText32
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText32");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText32", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText33
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText33");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText33", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText34
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText34");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText34", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText35
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText35");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText35", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText36
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText36");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText36", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText37
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText37");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText37", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText38
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText38");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText38", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText39
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText39");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText39", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseText40
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseText40");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseText40", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline1DurationEstimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline1DurationEstimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline1DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline2DurationEstimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline2DurationEstimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline2DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline3DurationEstimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline3DurationEstimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline3DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline4DurationEstimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline4DurationEstimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline4DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline5DurationEstimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline5DurationEstimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline5DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline6DurationEstimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline6DurationEstimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline6DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline7DurationEstimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline7DurationEstimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline7DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline8DurationEstimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline8DurationEstimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline8DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline9DurationEstimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline9DurationEstimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline9DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline10DurationEstimated
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline10DurationEstimated");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline10DurationEstimated", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object PhysicalPercentComplete
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PhysicalPercentComplete");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PhysicalPercentComplete", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjEarnedValueMethod EarnedValueMethod
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjEarnedValueMethod>(this, "EarnedValueMethod");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "EarnedValueMethod", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectCost1
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost1");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectCost1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectCost2
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost2");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectCost2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectCost3
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost3");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectCost3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectCost4
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost4");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectCost4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectCost5
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost5");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectCost5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectCost6
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost6");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectCost6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectCost7
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost7");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectCost7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectCost8
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost8");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectCost8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectCost9
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost9");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectCost9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectCost10
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectCost10");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectCost10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate1
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate1");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate2
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate2");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate3
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate3");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate4
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate4");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate5
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate5");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate6
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate6");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate7
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate7");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate8
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate8");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate9
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate9");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate10
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate10");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate11
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate11");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate12
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate12");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate13
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate13");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate14
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate14");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate15
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate15");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate16
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate16");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate17
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate17");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate18
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate18");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate19
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate19");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate20
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate20");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate21
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate21");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate22
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate22");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate23
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate23");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate24
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate24");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate25
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate25");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate26
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate26");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate27
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate27");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate28
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate28");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate29
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate29");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDate30
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDate30");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDate30", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDuration1
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration1");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDuration2
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration2");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDuration3
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration3");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDuration4
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration4");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDuration5
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration5");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDuration6
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration6");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDuration7
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration7");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDuration8
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration8");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDuration9
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration9");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectDuration10
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectDuration10");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectDuration10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode1
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode1");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode2
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode2");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode3
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode3");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode4
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode4");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode5
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode5");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode6
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode6");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode7
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode7");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode8
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode8");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode9
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode9");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode10
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode10");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode11
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode11");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode12
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode12");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode13
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode13");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode14
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode14");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode15
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode15");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode16
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode16");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode17
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode17");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode18
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode18");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode19
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode19");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode20
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode20");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode21
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode21");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode22
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode22");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode23
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode23");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode24
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode24");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode25
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode25");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode26
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode26");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode27
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode27");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode28
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode28");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode29
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode29");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectOutlineCode30
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectOutlineCode30");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectOutlineCode30", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag1
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag1");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag2
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag2");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag3
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag3");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag4
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag4");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag5
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag5");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag6
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag6");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag7
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag7");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag8
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag8");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag9
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag9");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag10
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag10");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag11
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag11");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag12
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag12");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag13
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag13");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag14
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag14");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag15
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag15");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag16
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag16");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag17
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag17");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag18
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag18");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag19
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag19");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object EnterpriseProjectFlag20
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnterpriseProjectFlag20");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "EnterpriseProjectFlag20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber1
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber1");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber2
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber2");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber3
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber3");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber4
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber4");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber5
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber5");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber6
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber6");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber7
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber7");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber8
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber8");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber9
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber9");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber10
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber10");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber11
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber11");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber12
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber12");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber13
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber13");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber14
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber14");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber15
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber15");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber16
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber16");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber17
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber17");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber18
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber18");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber19
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber19");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber20
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber20");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber21
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber21");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber22
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber22");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber23
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber23");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber24
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber24");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber25
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber25");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber26
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber26");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber27
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber27");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber28
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber28");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber29
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber29");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber30
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber30");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber30", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber31
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber31");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber31", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber32
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber32");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber32", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber33
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber33");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber33", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber34
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber34");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber34", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber35
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber35");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber35", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber36
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber36");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber36", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber37
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber37");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber37", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber38
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber38");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber38", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber39
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber39");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber39", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Double EnterpriseProjectNumber40
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "EnterpriseProjectNumber40");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectNumber40", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText1
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText1");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText2
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText2");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText3
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText3");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText4
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText4");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText5
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText5");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText6
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText6");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText7
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText7");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText8
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText8");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText9
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText9");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText10
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText10");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText11
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText11");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText12
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText12");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText13
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText13");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText14
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText14");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText15
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText15");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText16
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText16");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText17
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText17");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText18
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText18");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText19
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText19");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText20
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText20");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText21
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText21");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText22
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText22");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText23
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText23");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText24
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText24");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText25
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText25");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText26
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText26");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText27
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText27");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText28
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText28");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText29
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText29");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText30
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText30");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText30", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText31
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText31");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText31", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText32
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText32");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText32", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText33
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText33");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText33", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText34
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText34");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText34", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText35
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText35");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText35", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText36
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText36");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText36", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText37
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText37");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText37", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText38
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText38");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText38", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText39
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText39");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText39", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseProjectText40
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseProjectText40");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseProjectText40", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object ActualWorkProtected
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ActualWorkProtected");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ActualWorkProtected", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object ActualOvertimeWorkProtected
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ActualOvertimeWorkProtected");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ActualOvertimeWorkProtected", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object BaselineFixedCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BaselineFixedCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BaselineFixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline1FixedCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline1FixedCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline1FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline2FixedCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline2FixedCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline2FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline3FixedCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline3FixedCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline3FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline4FixedCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline4FixedCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline4FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline5FixedCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline5FixedCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline5FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline6FixedCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline6FixedCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline6FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline7FixedCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline7FixedCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline7FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline8FixedCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline8FixedCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline8FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline9FixedCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline9FixedCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline9FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline10FixedCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline10FixedCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline10FixedCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Guid
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Guid");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string CalendarGuid
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CalendarGuid");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string DeliverableGuid
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DeliverableGuid");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DeliverableGuid", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int16 DeliverableType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "DeliverableType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DeliverableType", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object DeliverableStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DeliverableStart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object DeliverableFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DeliverableFinish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object IsPublished
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "IsPublished");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "IsPublished", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string StatusManagerName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "StatusManagerName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StatusManagerName", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string ErrorMessage
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ErrorMessage");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object BudgetWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BudgetWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object BudgetCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BudgetCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjAccrueAt BaselineFixedCostAccrual
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "BaselineFixedCostAccrual");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BaselineFixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object BaselineDeliverableStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BaselineDeliverableStart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BaselineDeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object BaselineDeliverableFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BaselineDeliverableFinish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BaselineDeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object BaselineBudgetWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BaselineBudgetWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BaselineBudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object BaselineBudgetCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BaselineBudgetCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BaselineBudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline1FixedCostAccrual
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline1FixedCostAccrual");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Baseline1FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline1DeliverableStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline1DeliverableStart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline1DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline1DeliverableFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline1DeliverableFinish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline1DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline1BudgetWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline1BudgetWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline1BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline1BudgetCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline1BudgetCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline1BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline2FixedCostAccrual
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline2FixedCostAccrual");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Baseline2FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline2DeliverableStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline2DeliverableStart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline2DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline2DeliverableFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline2DeliverableFinish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline2DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline2BudgetWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline2BudgetWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline2BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline2BudgetCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline2BudgetCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline2BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline3FixedCostAccrual
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline3FixedCostAccrual");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Baseline3FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline3DeliverableStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline3DeliverableStart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline3DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline3DeliverableFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline3DeliverableFinish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline3DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline3BudgetWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline3BudgetWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline3BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline3BudgetCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline3BudgetCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline3BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline4FixedCostAccrual
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline4FixedCostAccrual");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Baseline4FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline4DeliverableStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline4DeliverableStart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline4DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline4DeliverableFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline4DeliverableFinish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline4DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline4BudgetWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline4BudgetWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline4BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline4BudgetCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline4BudgetCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline4BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline5FixedCostAccrual
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline5FixedCostAccrual");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Baseline5FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline5DeliverableStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline5DeliverableStart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline5DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline5DeliverableFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline5DeliverableFinish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline5DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline5BudgetWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline5BudgetWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline5BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline5BudgetCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline5BudgetCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline5BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline6FixedCostAccrual
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline6FixedCostAccrual");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Baseline6FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline6DeliverableStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline6DeliverableStart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline6DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline6DeliverableFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline6DeliverableFinish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline6DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline6BudgetWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline6BudgetWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline6BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline6BudgetCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline6BudgetCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline6BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline7FixedCostAccrual
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline7FixedCostAccrual");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Baseline7FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline7DeliverableStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline7DeliverableStart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline7DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline7DeliverableFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline7DeliverableFinish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline7DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline7BudgetWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline7BudgetWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline7BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline7BudgetCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline7BudgetCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline7BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline8FixedCostAccrual
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline8FixedCostAccrual");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Baseline8FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline8DeliverableStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline8DeliverableStart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline8DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline8DeliverableFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline8DeliverableFinish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline8DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline8BudgetWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline8BudgetWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline8BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline8BudgetCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline8BudgetCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline8BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline9FixedCostAccrual
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline9FixedCostAccrual");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Baseline9FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline9DeliverableStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline9DeliverableStart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline9DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline9DeliverableFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline9DeliverableFinish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline9DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline9BudgetWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline9BudgetWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline9BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline9BudgetCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline9BudgetCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline9BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjAccrueAt Baseline10FixedCostAccrual
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjAccrueAt>(this, "Baseline10FixedCostAccrual");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Baseline10FixedCostAccrual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline10DeliverableStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline10DeliverableStart");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline10DeliverableStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline10DeliverableFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline10DeliverableFinish");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline10DeliverableFinish", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline10BudgetWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline10BudgetWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline10BudgetWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Baseline10BudgetCost
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Baseline10BudgetCost");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Baseline10BudgetCost", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 RecalcFlags
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RecalcFlags");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.StartDriver StartDriver
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.StartDriver>(this, "StartDriver", NetOffice.MSProjectApi.StartDriver.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string DeliverableName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DeliverableName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DeliverableName", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public object Active
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Active");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Active", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public object Manual
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Manual");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Manual", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public object Placeholder
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Placeholder");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public object Warning
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Warning");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string StartText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "StartText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string FinishText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FinishText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string DurationText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DurationText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864567(v=office.14).aspx </remarks>
		[SupportByVersion("MSProject", 11,14)]
		public object IsStartValid
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "IsStartValid");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861711(v=office.14).aspx </remarks>
		[SupportByVersion("MSProject", 11,14)]
		public object IsFinishValid
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "IsFinishValid");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865936(v=office.14).aspx </remarks>
		[SupportByVersion("MSProject", 11,14)]
		public object IsDurationValid
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "IsDurationValid");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string BaselineStartText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BaselineStartText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BaselineStartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string BaselineFinishText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BaselineFinishText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BaselineFinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string BaselineDurationText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BaselineDurationText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BaselineDurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline1StartText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline1StartText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline1StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline1FinishText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline1FinishText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline1FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline1DurationText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline1DurationText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline1DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline2StartText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline2StartText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline2StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline2FinishText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline2FinishText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline2FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline2DurationText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline2DurationText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline2DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline3StartText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline3StartText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline3StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline3FinishText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline3FinishText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline3FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline3DurationText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline3DurationText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline3DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline4StartText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline4StartText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline4StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline4FinishText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline4FinishText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline4FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline4DurationText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline4DurationText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline4DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline5StartText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline5StartText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline5StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline5FinishText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline5FinishText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline5FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline5DurationText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline5DurationText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline5DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline6StartText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline6StartText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline6StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline6FinishText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline6FinishText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline6FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline6DurationText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline6DurationText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline6DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline7StartText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline7StartText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline7StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline7FinishText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline7FinishText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline7FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline7DurationText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline7DurationText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline7DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline8StartText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline8StartText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline8StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline8FinishText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline8FinishText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline8FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline8DurationText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline8DurationText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline8DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline9StartText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline9StartText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline9StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline9FinishText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline9FinishText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline9FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline9DurationText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline9DurationText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline9DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline10StartText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline10StartText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline10StartText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline10FinishText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline10FinishText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline10FinishText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public string Baseline10DurationText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Baseline10DurationText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Baseline10DurationText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public object IgnoreWarnings
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "IgnoreWarnings");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "IgnoreWarnings", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public NetOffice.MSProjectApi.Calendar CalendarObject
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Calendar>(this, "CalendarObject", NetOffice.MSProjectApi.Calendar.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public object ScheduledStart
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ScheduledStart");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public object ScheduledFinish
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ScheduledFinish");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public object ScheduledDuration
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ScheduledDuration");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public object PathDrivingPredecessor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PathDrivingPredecessor");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public object PathPredecessor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PathPredecessor");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public object PathDrivenSuccessor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PathDrivenSuccessor");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public object PathSuccessor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PathSuccessor");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="fieldID">NetOffice.MSProjectApi.Enums.PjField fieldID</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public string GetField(NetOffice.MSProjectApi.Enums.PjField fieldID)
		{
			return Factory.ExecuteStringMethodGet(this, "GetField", fieldID);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="fieldID">NetOffice.MSProjectApi.Enums.PjField fieldID</param>
		/// <param name="value">string value</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void SetField(NetOffice.MSProjectApi.Enums.PjField fieldID, string value)
		{
			 Factory.ExecuteMethod(this, "SetField", fieldID, value);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
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
		/// <param name="startDate">object startDate</param>
		/// <param name="endDate">object endDate</param>
		/// <param name="type">optional NetOffice.MSProjectApi.Enums.PjTaskTimescaledData Type = 0</param>
		/// <param name="timeScaleUnit">optional NetOffice.MSProjectApi.Enums.PjTimescaleUnit TimeScaleUnit = 3</param>
		/// <param name="count">optional Int32 Count = 1</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TimeScaleValues TimeScaleData(object startDate, object endDate, object type, object timeScaleUnit, object count)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TimeScaleValues>(this, "TimeScaleData", NetOffice.MSProjectApi.TimeScaleValues.LateBindingApiWrapperType, new object[]{ startDate, endDate, type, timeScaleUnit, count });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="startDate">object startDate</param>
		/// <param name="endDate">object endDate</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TimeScaleValues TimeScaleData(object startDate, object endDate)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TimeScaleValues>(this, "TimeScaleData", NetOffice.MSProjectApi.TimeScaleValues.LateBindingApiWrapperType, startDate, endDate);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="startDate">object startDate</param>
		/// <param name="endDate">object endDate</param>
		/// <param name="type">optional NetOffice.MSProjectApi.Enums.PjTaskTimescaledData Type = 0</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.TimeScaleValues TimeScaleData(object startDate, object endDate, object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TimeScaleValues>(this, "TimeScaleData", NetOffice.MSProjectApi.TimeScaleValues.LateBindingApiWrapperType, startDate, endDate, type);
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
		public NetOffice.MSProjectApi.TimeScaleValues TimeScaleData(object startDate, object endDate, object type, object timeScaleUnit)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.TimeScaleValues>(this, "TimeScaleData", NetOffice.MSProjectApi.TimeScaleValues.LateBindingApiWrapperType, startDate, endDate, type, timeScaleUnit);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tasks">object tasks</param>
		/// <param name="link">optional NetOffice.MSProjectApi.Enums.PjTaskLinkType Link = 1</param>
		/// <param name="lag">optional object lag</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void LinkPredecessors(object tasks, object link, object lag)
		{
			 Factory.ExecuteMethod(this, "LinkPredecessors", tasks, link, lag);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tasks">object tasks</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public void LinkPredecessors(object tasks)
		{
			 Factory.ExecuteMethod(this, "LinkPredecessors", tasks);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tasks">object tasks</param>
		/// <param name="link">optional NetOffice.MSProjectApi.Enums.PjTaskLinkType Link = 1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public void LinkPredecessors(object tasks, object link)
		{
			 Factory.ExecuteMethod(this, "LinkPredecessors", tasks, link);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tasks">object tasks</param>
		/// <param name="link">optional NetOffice.MSProjectApi.Enums.PjTaskLinkType Link = 1</param>
		/// <param name="lag">optional object lag</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void LinkSuccessors(object tasks, object link, object lag)
		{
			 Factory.ExecuteMethod(this, "LinkSuccessors", tasks, link, lag);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tasks">object tasks</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public void LinkSuccessors(object tasks)
		{
			 Factory.ExecuteMethod(this, "LinkSuccessors", tasks);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tasks">object tasks</param>
		/// <param name="link">optional NetOffice.MSProjectApi.Enums.PjTaskLinkType Link = 1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public void LinkSuccessors(object tasks, object link)
		{
			 Factory.ExecuteMethod(this, "LinkSuccessors", tasks, link);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tasks">object tasks</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void UnlinkPredecessors(object tasks)
		{
			 Factory.ExecuteMethod(this, "UnlinkPredecessors", tasks);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="tasks">object tasks</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void UnlinkSuccessors(object tasks)
		{
			 Factory.ExecuteMethod(this, "UnlinkSuccessors", tasks);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public void OutlineIndent()
		{
			 Factory.ExecuteMethod(this, "OutlineIndent");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public void OutlineOutdent()
		{
			 Factory.ExecuteMethod(this, "OutlineOutdent");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public void OutlineHideSubTasks()
		{
			 Factory.ExecuteMethod(this, "OutlineHideSubTasks");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public void OutlineShowSubTasks()
		{
			 Factory.ExecuteMethod(this, "OutlineShowSubTasks");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public void OutlineShowAllTasks()
		{
			 Factory.ExecuteMethod(this, "OutlineShowAllTasks");
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="startSplitOn">object startSplitOn</param>
		/// <param name="endSplitOn">object endSplitOn</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public void Split(object startSplitOn, object endSplitOn)
		{
			 Factory.ExecuteMethod(this, "Split", startSplitOn, endSplitOn);
		}

		#endregion

		#pragma warning restore
	}
}
