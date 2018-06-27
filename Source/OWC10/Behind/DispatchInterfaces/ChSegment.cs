using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface ChSegment 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ChSegment : COMObject, NetOffice.OWC10Api.ChSegment
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
                    _contractType = typeof(NetOffice.OWC10Api.ChSegment);
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
                    _type = typeof(ChSegment);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ChSegment() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.ChSegmentBoundary Begin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChSegmentBoundary>(this, "Begin", typeof(NetOffice.OWC10Api.ChSegmentBoundary));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.ChSegmentBoundary End
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChSegmentBoundary>(this, "End", typeof(NetOffice.OWC10Api.ChSegmentBoundary));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool HasAbsoluteLabels
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasAbsoluteLabels");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasAbsoluteLabels", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool HasDiscreteDivisions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasDiscreteDivisions");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasDiscreteDivisions", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool HasAutoDivisions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasAutoDivisions");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasAutoDivisions", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Divisions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Divisions");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Divisions", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}


