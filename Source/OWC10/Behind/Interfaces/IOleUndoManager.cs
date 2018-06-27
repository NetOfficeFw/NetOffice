using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// Interface IOleUndoManager 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
 	public class IOleUndoManager : COMObject, NetOffice.OWC10Api.IOleUndoManager
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
                    _contractType = typeof(NetOffice.OWC10Api.IOleUndoManager);
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
                    _type = typeof(IOleUndoManager);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IOleUndoManager() : base()
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
		/// <param name="pdwState">Int32 pdwState</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 GetOpenParentState(out Int32 pdwState)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pdwState = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pdwState);
			object returnItem = Invoker.MethodReturn(this, "GetOpenParentState", paramsArray, modifiers);
			pdwState = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pUU">NetOffice.OWC10Api.IOleUndoUnit pUU</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 DiscardFrom(NetOffice.OWC10Api.IOleUndoUnit pUU)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DiscardFrom", pUU);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pUU">NetOffice.OWC10Api.IOleUndoUnit pUU</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 UndoTo(NetOffice.OWC10Api.IOleUndoUnit pUU)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "UndoTo", pUU);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pUU">NetOffice.OWC10Api.IOleUndoUnit pUU</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 RedoTo(NetOffice.OWC10Api.IOleUndoUnit pUU)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RedoTo", pUU);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="ppEnum">NetOffice.OWC10Api.IEnumOleUndoUnits ppEnum</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 EnumUndoable(out NetOffice.OWC10Api.IEnumOleUndoUnits ppEnum)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppEnum = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppEnum);
			object returnItem = Invoker.MethodReturn(this, "EnumUndoable", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppEnum = Factory.CreateObjectFromComProxy(this, paramsArray[0], false) as NetOffice.OWC10Api.IEnumOleUndoUnits;
            else
                ppEnum = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="ppEnum">NetOffice.OWC10Api.IEnumOleUndoUnits ppEnum</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 EnumRedoable(out NetOffice.OWC10Api.IEnumOleUndoUnits ppEnum)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppEnum = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppEnum);
			object returnItem = Invoker.MethodReturn(this, "EnumRedoable", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppEnum = Factory.CreateObjectFromComProxy(this, paramsArray[0], false) as NetOffice.OWC10Api.IEnumOleUndoUnits;
            else
                ppEnum = null;
            return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pbstr">string pbstr</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 GetLastUndoDescription(out string pbstr)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pbstr = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(pbstr);
			object returnItem = Invoker.MethodReturn(this, "GetLastUndoDescription", paramsArray, modifiers);
			pbstr = paramsArray[0] as string;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pbstr">string pbstr</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 GetLastRedoDescription(out string pbstr)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pbstr = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(pbstr);
			object returnItem = Invoker.MethodReturn(this, "GetLastRedoDescription", paramsArray, modifiers);
			pbstr = paramsArray[0] as string;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="fEnable">Int32 fEnable</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Enable(Int32 fEnable)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Enable", fEnable);
		}

		#endregion

		#pragma warning restore
	}
}

