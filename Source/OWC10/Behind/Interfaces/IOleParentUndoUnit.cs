using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// Interface IOleParentUndoUnit 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
 	public class IOleParentUndoUnit : IOleUndoUnit, NetOffice.OWC10Api.IOleParentUndoUnit
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
                    _contractType = typeof(NetOffice.OWC10Api.IOleParentUndoUnit);
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
                    _type = typeof(IOleParentUndoUnit);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IOleParentUndoUnit() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pPUU">NetOffice.OWC10Api.IOleParentUndoUnit pPUU</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Open(NetOffice.OWC10Api.IOleParentUndoUnit pPUU)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Open", pPUU);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pPUU">NetOffice.OWC10Api.IOleParentUndoUnit pPUU</param>
		/// <param name="fCommit">Int32 fCommit</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Close(NetOffice.OWC10Api.IOleParentUndoUnit pPUU, Int32 fCommit)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Close", pPUU, fCommit);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pUU">NetOffice.OWC10Api.IOleUndoUnit pUU</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Add(NetOffice.OWC10Api.IOleUndoUnit pUU)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Add", pUU);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pUU">NetOffice.OWC10Api.IOleUndoUnit pUU</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 FindUnit(NetOffice.OWC10Api.IOleUndoUnit pUU)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "FindUnit", pUU);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pdwState">Int32 pdwState</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 GetParentState(out Int32 pdwState)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pdwState = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pdwState);
			object returnItem = Invoker.MethodReturn(this, "GetParentState", paramsArray, modifiers);
			pdwState = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

