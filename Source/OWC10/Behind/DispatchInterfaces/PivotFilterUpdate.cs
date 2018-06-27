using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface PivotFilterUpdate 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PivotFilterUpdate : COMObject, NetOffice.OWC10Api.PivotFilterUpdate
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
                    _contractType = typeof(NetOffice.OWC10Api.PivotFilterUpdate);
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
                    _type = typeof(PivotFilterUpdate);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public PivotFilterUpdate() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="member">NetOffice.OWC10Api.PivotMember member</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum get_StateOf(NetOffice.OWC10Api.PivotMember member)
		{
			return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum>(this, "StateOf", member);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_StateOf
		/// </summary>
		/// <param name="member">NetOffice.OWC10Api.PivotMember member</param>
		[SupportByVersion("OWC10", 1), Redirect("get_StateOf")]
		public virtual NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum StateOf(NetOffice.OWC10Api.PivotMember member)
		{
			return get_StateOf(member);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool IsDirty
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsDirty");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="member">NetOffice.OWC10Api.PivotMember member</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void Click(NetOffice.OWC10Api.PivotMember member)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Click", member);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="member">NetOffice.OWC10Api.PivotMember member</param>
		/// <param name="oldMemberState">NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum oldMemberState</param>
		/// <param name="newMemberState">NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum newMemberState</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void ClickFromTo(NetOffice.OWC10Api.PivotMember member, NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum oldMemberState, NetOffice.OWC10Api.Enums.PivotFilterUpdateMemberStateEnum newMemberState)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ClickFromTo", member, oldMemberState, newMemberState);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void Apply()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Apply");
		}

		#endregion

		#pragma warning restore
	}
}

