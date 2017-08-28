using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface Assignment 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920546(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Assignment : COMObject
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
                    _type = typeof(Assignment);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Assignment(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Assignment(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Assignment(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Assignment(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Assignment(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Assignment(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Assignment() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Assignment(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

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
		public Int32 TaskUniqueID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TaskUniqueID");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TaskUniqueID", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 ResourceUniqueID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ResourceUniqueID");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ResourceUniqueID", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 TaskID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TaskID");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TaskID", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public Int32 ResourceID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ResourceID");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ResourceID", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string TaskName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TaskName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TaskName", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string ResourceName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ResourceName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ResourceName", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Units
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Units");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Units", value);
			}
		}

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
		public object OvertimeWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "OvertimeWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "OvertimeWork", value);
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object RegularWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "RegularWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "RegularWork", value);
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
		public object ActualOvertimeWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ActualOvertimeWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ActualOvertimeWork", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object RemainingOvertimeWork
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "RemainingOvertimeWork");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "RemainingOvertimeWork", value);
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
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Peak
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Peak");
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
		public object ACWP
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ACWP");
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjWorkContourType WorkContour
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjWorkContourType>(this, "WorkContour");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "WorkContour", value);
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
		public object Confirmed
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Confirmed");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Confirmed", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object ResponsePending
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ResponsePending");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ResponsePending", value);
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
		public object TeamStatusPending
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TeamStatusPending");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "TeamStatusPending", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object CostRateTable
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CostRateTable");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "CostRateTable", value);
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
		public string TaskSummaryName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TaskSummaryName");
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
		public object FixedMaterialAssignment
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FixedMaterialAssignment");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjResourceTypes ResourceType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjResourceTypes>(this, "ResourceType");
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
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string TaskOutlineNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TaskOutlineNumber");
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
		public string EnterpriseResourceOutlineCode1
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode1");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode1", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode2
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode2");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode2", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode3
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode3");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode3", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode4
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode4");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode4", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode5
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode5");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode5", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode6
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode6");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode6", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode7
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode7");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode7", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode8
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode8");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode8", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode9
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode9");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode9", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode10
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode10");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode10", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode11
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode11");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode11", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode12
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode12");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode12", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode13
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode13");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode13", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode14
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode14");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode14", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode15
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode15");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode15", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode16
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode16");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode16", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode17
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode17");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode17", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode18
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode18");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode18", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode19
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode19");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode19", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode20
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode20");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode21
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode21");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode22
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode22");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode23
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode23");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode24
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode24");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode25
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode25");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode26
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode26");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode27
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode27");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode28
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode28");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceOutlineCode29
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceOutlineCode29");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceOutlineCode29", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceRBS
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceRBS");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceRBS", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjResourceRequestType ResourceRequestType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjResourceRequestType>(this, "ResourceRequestType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ResourceRequestType", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceMultiValue20
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceMultiValue20");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceMultiValue20", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceMultiValue21
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceMultiValue21");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceMultiValue21", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceMultiValue22
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceMultiValue22");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceMultiValue22", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceMultiValue23
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceMultiValue23");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceMultiValue23", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceMultiValue24
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceMultiValue24");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceMultiValue24", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceMultiValue25
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceMultiValue25");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceMultiValue25", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceMultiValue26
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceMultiValue26");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceMultiValue26", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceMultiValue27
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceMultiValue27");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceMultiValue27", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceMultiValue28
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceMultiValue28");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceMultiValue28", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string EnterpriseResourceMultiValue29
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EnterpriseResourceMultiValue29");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnterpriseResourceMultiValue29", value);
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
		public object ActualOvtWorkProtected
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ActualOvtWorkProtected");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ActualOvtWorkProtected", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public object Created
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Created");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Created", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string WBS
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "WBS");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Enums.PjBookingTypes BookingType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSProjectApi.Enums.PjBookingTypes>(this, "BookingType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BookingType", value);
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
		public string TaskGuid
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TaskGuid");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string ResourceGuid
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ResourceGuid");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Summary
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Summary");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public string Owner
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Owner");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Owner", value);
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
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public NetOffice.MSProjectApi.Task Task
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Task>(this, "Task", NetOffice.MSProjectApi.Task.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		public NetOffice.MSProjectApi.Resource Resource
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Resource>(this, "Resource", NetOffice.MSProjectApi.Resource.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

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
		/// <param name="type">optional NetOffice.MSProjectApi.Enums.PjAssignmentTimescaledData Type = 8</param>
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
		/// <param name="type">optional NetOffice.MSProjectApi.Enums.PjAssignmentTimescaledData Type = 8</param>
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
		/// <param name="type">optional NetOffice.MSProjectApi.Enums.PjAssignmentTimescaledData Type = 8</param>
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
		/// <param name="project">object project</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public bool EnterpriseTeamMember(object project)
		{
			return Factory.ExecuteBoolMethodGet(this, "EnterpriseTeamMember", project);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		/// <param name="action">NetOffice.MSProjectApi.Enums.PjAssignmentReplanAction action</param>
		[SupportByVersion("MSProject", 11,14)]
		public void Replan(NetOffice.MSProjectApi.Enums.PjAssignmentReplanAction action)
		{
			 Factory.ExecuteMethod(this, "Replan", action);
		}

		#endregion

		#pragma warning restore
	}
}
