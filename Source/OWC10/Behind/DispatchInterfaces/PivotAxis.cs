using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface PivotAxis 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class PivotAxis : COMObject, NetOffice.OWC10Api.PivotAxis
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
                    _contractType = typeof(NetOffice.OWC10Api.PivotAxis);
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
                    _type = typeof(PivotAxis);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public PivotAxis() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotView View
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotView>(this, "View", typeof(NetOffice.OWC10Api.PivotView));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotFieldSets FieldSets
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotFieldSets>(this, "FieldSets", typeof(NetOffice.OWC10Api.PivotFieldSets));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.PivotLabel Label
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotLabel>(this, "Label", typeof(NetOffice.OWC10Api.PivotLabel));
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="fieldSet">NetOffice.OWC10Api.PivotFieldSet fieldSet</param>
		/// <param name="before">optional object before</param>
		/// <param name="remove">optional bool Remove = true</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void InsertFieldSet(NetOffice.OWC10Api.PivotFieldSet fieldSet, object before, object remove)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "InsertFieldSet", fieldSet, before, remove);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="fieldSet">NetOffice.OWC10Api.PivotFieldSet fieldSet</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void InsertFieldSet(NetOffice.OWC10Api.PivotFieldSet fieldSet)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "InsertFieldSet", fieldSet);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="fieldSet">NetOffice.OWC10Api.PivotFieldSet fieldSet</param>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public virtual void InsertFieldSet(NetOffice.OWC10Api.PivotFieldSet fieldSet, object before)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "InsertFieldSet", fieldSet, before);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="fieldSet">object fieldSet</param>
		[SupportByVersion("OWC10", 1)]
		public virtual void RemoveFieldSet(object fieldSet)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveFieldSet", fieldSet);
		}

		#endregion

		#pragma warning restore
	}
}


