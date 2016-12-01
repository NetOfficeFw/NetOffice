using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OWC10Api
{
	///<summary>
	/// Interface FieldListHierarchySite 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class FieldListHierarchySite : COMObject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public FieldListHierarchySite(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FieldListHierarchySite(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FieldListHierarchySite(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FieldListHierarchySite(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FieldListHierarchySite(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FieldListHierarchySite() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FieldListHierarchySite(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="nOldNodeId">Int32 nOldNodeId</param>
		/// <param name="nOldTypeId">Int32 nOldTypeId</param>
		/// <param name="pfPrevent">Int32 pfPrevent</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 PreSelect(Int32 nNodeId, Int32 nTypeId, Int32 nOldNodeId, Int32 nOldTypeId, out Int32 pfPrevent)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true);
			pfPrevent = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(nNodeId, nTypeId, nOldNodeId, nOldTypeId, pfPrevent);
			object returnItem = Invoker.MethodReturn(this, "PreSelect", paramsArray);
			pfPrevent = (Int32)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="nOldNodeId">Int32 nOldNodeId</param>
		/// <param name="nOldTypeId">Int32 nOldTypeId</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 PostSelect(Int32 nNodeId, Int32 nTypeId, Int32 nOldNodeId, Int32 nOldTypeId)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nNodeId, nTypeId, nOldNodeId, nOldTypeId);
			object returnItem = Invoker.MethodReturn(this, "PostSelect", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="fExpand">Int32 fExpand</param>
		/// <param name="pfPrevent">Int32 pfPrevent</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 PreExpand(Int32 nNodeId, Int32 nTypeId, Int32 fExpand, out Int32 pfPrevent)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true);
			pfPrevent = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(nNodeId, nTypeId, fExpand, pfPrevent);
			object returnItem = Invoker.MethodReturn(this, "PreExpand", paramsArray);
			pfPrevent = (Int32)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="fExpand">Int32 fExpand</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 PostExpand(Int32 nNodeId, Int32 nTypeId, Int32 fExpand)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nNodeId, nTypeId, fExpand);
			object returnItem = Invoker.MethodReturn(this, "PostExpand", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="ppobject">object ppobject</param>
		/// <param name="ppPivotView">object ppPivotView</param>
		/// <param name="pfPrevent">Int32 pfPrevent</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 PreDrag(Int32 nNodeId, Int32 nTypeId, out object ppobject, out object ppPivotView, out Int32 pfPrevent)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true,true);
			ppobject = null;
			ppPivotView = null;
			pfPrevent = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(nNodeId, nTypeId, ppobject, ppPivotView, pfPrevent);
			object returnItem = Invoker.MethodReturn(this, "PreDrag", paramsArray);
			ppobject = (object)paramsArray[2];
			ppPivotView = (object)paramsArray[3];
			pfPrevent = (Int32)paramsArray[4];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="hRes">Int32 hRes</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 PostDrag(Int32 nNodeId, Int32 nTypeId, Int32 hRes)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nNodeId, nTypeId, hRes);
			object returnItem = Invoker.MethodReturn(this, "PostDrag", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 PopulateChildren(Int32 nNodeId, Int32 nTypeId)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nNodeId, nTypeId);
			object returnItem = Invoker.MethodReturn(this, "PopulateChildren", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="hMenu">UIntPtr hMenu</param>
		/// <param name="pfPrevent">Int32 pfPrevent</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 ContextMenu(Int32 nNodeId, Int32 nTypeId, UIntPtr hMenu, out Int32 pfPrevent)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true);
			pfPrevent = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(nNodeId, nTypeId, hMenu, pfPrevent);
			object returnItem = Invoker.MethodReturn(this, "ContextMenu", paramsArray);
			pfPrevent = (Int32)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="wid">UIntPtr wid</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 DoCommand(Int32 nNodeId, Int32 nTypeId, UIntPtr wid)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nNodeId, nTypeId, wid);
			object returnItem = Invoker.MethodReturn(this, "DoCommand", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 DoubleClick(Int32 nNodeId, Int32 nTypeId)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nNodeId, nTypeId);
			object returnItem = Invoker.MethodReturn(this, "DoubleClick", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 PostDelete(Int32 nNodeId, Int32 nTypeId)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nNodeId, nTypeId);
			object returnItem = Invoker.MethodReturn(this, "PostDelete", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="nSelMask">Int32 nSelMask</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 PostMSelect(Int32 nSelMask)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nSelMask);
			object returnItem = Invoker.MethodReturn(this, "PostMSelect", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 Click(Int32 nNodeId, Int32 nTypeId)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nNodeId, nTypeId);
			object returnItem = Invoker.MethodReturn(this, "Click", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="nNodeId">Int32 nNodeId</param>
		/// <param name="nTypeId">Int32 nTypeId</param>
		/// <param name="nMsg">Int32 nMsg</param>
		/// <param name="nwParam">Int32 nwParam</param>
		/// <param name="nlParam">Int32 nlParam</param>
		/// <param name="pfStopProcessing">Int32 pfStopProcessing</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 KeyEvent(Int32 nNodeId, Int32 nTypeId, Int32 nMsg, Int32 nwParam, Int32 nlParam, out Int32 pfStopProcessing)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,true);
			pfStopProcessing = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(nNodeId, nTypeId, nMsg, nwParam, nlParam, pfStopProcessing);
			object returnItem = Invoker.MethodReturn(this, "KeyEvent", paramsArray);
			pfStopProcessing = (Int32)paramsArray[5];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}