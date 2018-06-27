using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface PivotDataAxis 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PivotDataAxis : PivotAxis, NetOffice.OWC10Api.PivotDataAxis
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
                    _contractType = typeof(NetOffice.OWC10Api.PivotDataAxis);
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
                    _type = typeof(PivotDataAxis);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public PivotDataAxis() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotTotals Totals
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotTotals>(this, "Totals", typeof(NetOffice.OWC10Api.PivotTotals));
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="total">NetOffice.OWC10Api.PivotTotal total</param>
		/// <param name="before">optional object before</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void InsertTotal(NetOffice.OWC10Api.PivotTotal total, object before)		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "InsertTotal", total, before);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="total">NetOffice.OWC10Api.PivotTotal total</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void InsertTotal(NetOffice.OWC10Api.PivotTotal total)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "InsertTotal", total);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="total">object total</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void RemoveTotal(object total)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveTotal", total);
		}

		#endregion

		#pragma warning restore
	}
}


