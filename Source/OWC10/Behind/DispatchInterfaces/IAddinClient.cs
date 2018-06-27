using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface IAddinClient 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IAddinClient : COMObject, NetOffice.OWC10Api.IAddinClient
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
                    _contractType = typeof(NetOffice.OWC10Api.IAddinClient);
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
                    _type = typeof(IAddinClient);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IAddinClient() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="vardisp">object vardisp</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void GrantAddinHost(object vardisp)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "GrantAddinHost", vardisp);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void RemoveAddinHost()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveAddinHost");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		/// <param name="semiCalced">bool semiCalced</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void IsSemiCalced(Int32 dispid, bool semiCalced)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "IsSemiCalced", dispid, semiCalced);
		}

		#endregion

		#pragma warning restore
	}
}

