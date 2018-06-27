using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// Interface FieldListHierarchySite 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
 	public class FieldListHierarchySite : COMObject, NetOffice.OWC10Api.FieldListHierarchySite
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
                    _contractType = typeof(NetOffice.OWC10Api.FieldListHierarchySite);
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
                    _type = typeof(FieldListHierarchySite);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public FieldListHierarchySite() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="nOldNodeId">Int32 nOldNodeId</param>
		/// <param name="nOldTypeId">Int32 nOldTypeId</param>
		/// <param name="pfPrevent">Int32 pfPrevent</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 PreSelect(Int32 nNodeId, Int32 nTypeId, Int32 nOldNodeId, Int32 nOldTypeId, out Int32 pfPrevent)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true);
			pfPrevent = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(nNodeId, nTypeId, nOldNodeId, nOldTypeId, pfPrevent);
			object returnItem = Invoker.MethodReturn(this, "PreSelect", paramsArray, modifiers);
			pfPrevent = (Int32)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="nOldNodeId">Int32 nOldNodeId</param>
		/// <param name="nOldTypeId">Int32 nOldTypeId</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 PostSelect(Int32 nNodeId, Int32 nTypeId, Int32 nOldNodeId, Int32 nOldTypeId)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PostSelect", nNodeId, nTypeId, nOldNodeId, nOldTypeId);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="fExpand">Int32 fExpand</param>
		/// <param name="pfPrevent">Int32 pfPrevent</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 PreExpand(Int32 nNodeId, Int32 nTypeId, Int32 fExpand, out Int32 pfPrevent)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true);
			pfPrevent = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(nNodeId, nTypeId, fExpand, pfPrevent);
			object returnItem = Invoker.MethodReturn(this, "PreExpand", paramsArray, modifiers);
			pfPrevent = (Int32)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="fExpand">Int32 fExpand</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 PostExpand(Int32 nNodeId, Int32 nTypeId, Int32 fExpand)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PostExpand", nNodeId, nTypeId, fExpand);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="ppobject">object ppobject</param>
		/// <param name="ppPivotView">object ppPivotView</param>
		/// <param name="pfPrevent">Int32 pfPrevent</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 PreDrag(Int32 nNodeId, Int32 nTypeId, out object ppobject, out object ppPivotView, out Int32 pfPrevent)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true,true);
			ppobject = null;
			ppPivotView = null;
			pfPrevent = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(nNodeId, nTypeId, ppobject, ppPivotView, pfPrevent);
			object returnItem = Invoker.MethodReturn(this, "PreDrag", paramsArray, modifiers);
			ppobject = (object)paramsArray[2];
			ppPivotView = (object)paramsArray[3];
			pfPrevent = (Int32)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="hRes">Int32 hRes</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 PostDrag(Int32 nNodeId, Int32 nTypeId, Int32 hRes)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PostDrag", nNodeId, nTypeId, hRes);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 PopulateChildren(Int32 nNodeId, Int32 nTypeId)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PopulateChildren", nNodeId, nTypeId);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="hMenu">UIntPtr hMenu</param>
		/// <param name="pfPrevent">Int32 pfPrevent</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 ContextMenu(Int32 nNodeId, Int32 nTypeId, UIntPtr hMenu, out Int32 pfPrevent)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true);
			pfPrevent = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(nNodeId, nTypeId, hMenu, pfPrevent);
			object returnItem = Invoker.MethodReturn(this, "ContextMenu", paramsArray, modifiers);
			pfPrevent = (Int32)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="wid">UIntPtr wid</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 DoCommand(Int32 nNodeId, Int32 nTypeId, UIntPtr wid)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DoCommand", nNodeId, nTypeId, wid);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 DoubleClick(Int32 nNodeId, Int32 nTypeId)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DoubleClick", nNodeId, nTypeId);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 PostDelete(Int32 nNodeId, Int32 nTypeId)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PostDelete", nNodeId, nTypeId);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nSelMask">Int32 nSelMask</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 PostMSelect(Int32 nSelMask)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PostMSelect", nSelMask);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Click(Int32 nNodeId, Int32 nTypeId)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Click", nNodeId, nTypeId);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="nMsg">Int32 nMsg</param>
		/// <param name="nwParam">Int32 nwParam</param>
		/// <param name="nlParam">Int32 nlParam</param>
		/// <param name="pfStopProcessing">Int32 pfStopProcessing</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 KeyEvent(Int32 nNodeId, Int32 nTypeId, Int32 nMsg, Int32 nwParam, Int32 nlParam, out Int32 pfStopProcessing)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,true);
			pfStopProcessing = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(nNodeId, nTypeId, nMsg, nwParam, nlParam, pfStopProcessing);
			object returnItem = Invoker.MethodReturn(this, "KeyEvent", paramsArray, modifiers);
			pfStopProcessing = (Int32)paramsArray[5];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

