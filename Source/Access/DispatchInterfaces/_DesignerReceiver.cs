using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.AccessApi
{
	///<summary>
	/// DispatchInterface _DesignerReceiver 
	/// SupportByVersion Access, 15, 16
	///</summary>
	[SupportByVersionAttribute("Access", 15, 16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _DesignerReceiver : COMObject
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
                    _type = typeof(_DesignerReceiver);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _DesignerReceiver(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _DesignerReceiver(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _DesignerReceiver(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _DesignerReceiver(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _DesignerReceiver(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _DesignerReceiver() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _DesignerReceiver(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void SetReady()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SetReady", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void RefreshRibbon()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RefreshRibbon", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="builder">Int16 builder</param>
		/// <param name="bstrBuilderValue">string bstrBuilderValue</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public object LaunchBuilder(Int16 builder, string bstrBuilderValue)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(builder, bstrBuilderValue);
			object returnItem = Invoker.MethodReturn(this, "LaunchBuilder", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="propertytype">Int16 propertytype</param>
		/// <param name="inputData">optional object inputData</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public object RetrievePropertyValues(Int16 propertytype, object inputData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(propertytype, inputData);
			object returnItem = Invoker.MethodReturn(this, "RetrievePropertyValues", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="propertytype">Int16 propertytype</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 15, 16)]
		public object RetrievePropertyValues(Int16 propertytype)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(propertytype);
			object returnItem = Invoker.MethodReturn(this, "RetrievePropertyValues", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void RecordSourceUpdated(string bstrRecordSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrRecordSource);
			Invoker.Method(this, "RecordSourceUpdated", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public string GetEmbeddedRecordSourceSQL()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetEmbeddedRecordSourceSQL", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public string SetEmbeddedRecordSourceSQL(string bstrRecordSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrRecordSource);
			object returnItem = Invoker.MethodReturn(this, "SetEmbeddedRecordSourceSQL", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="bstrViewName">string bstrViewName</param>
		/// <param name="viewtype">Int16 viewtype</param>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		/// <param name="fStandalone">bool fStandalone</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void QuickCreateView(string bstrViewName, Int16 viewtype, string bstrRecordSource, bool fStandalone)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrViewName, viewtype, bstrRecordSource, fStandalone);
			Invoker.Method(this, "QuickCreateView", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="bstrExpression">string bstrExpression</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public object IsExpressionValid(string bstrExpression)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrExpression);
			object returnItem = Invoker.MethodReturn(this, "IsExpressionValid", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="bstrOldCtrlID">string bstrOldCtrlID</param>
		/// <param name="bstrNewCtrlID">string bstrNewCtrlID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void NotifyControlIDChanged(string bstrOldCtrlID, string bstrNewCtrlID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrOldCtrlID, bstrNewCtrlID);
			Invoker.Method(this, "NotifyControlIDChanged", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public object RetrieveControlSourcesInfo(string bstrRecordSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrRecordSource);
			object returnItem = Invoker.MethodReturn(this, "RetrieveControlSourcesInfo", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="accessObjectType">Int32 accessObjectType</param>
		/// <param name="accessObjectName">string accessObjectName</param>
		/// <param name="fDesignMode">bool fDesignMode</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void OpenAccessObject(Int32 accessObjectType, string accessObjectName, bool fDesignMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(accessObjectType, accessObjectName, fDesignMode);
			Invoker.Method(this, "OpenAccessObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="accessObjectType">Int32 accessObjectType</param>
		/// <param name="accessObjectName">string accessObjectName</param>
		/// <param name="pdispDependentObjectTypeNamePairArray">object pdispDependentObjectTypeNamePairArray</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public bool DeleteAccessObject(Int32 accessObjectType, string accessObjectName, object pdispDependentObjectTypeNamePairArray)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(accessObjectType, accessObjectName, pdispDependentObjectTypeNamePairArray);
			object returnItem = Invoker.MethodReturn(this, "DeleteAccessObject", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="accessObjectType">Int32 accessObjectType</param>
		/// <param name="accessObjectName">string accessObjectName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void DuplicateAccessObject(Int32 accessObjectType, string accessObjectName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(accessObjectType, accessObjectName);
			Invoker.Method(this, "DuplicateAccessObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public object RetrieveRecordSourceInfo(string bstrRecordSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrRecordSource);
			object returnItem = Invoker.MethodReturn(this, "RetrieveRecordSourceInfo", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="bstrViewName">string bstrViewName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public object RetrieveViewInfo(string bstrViewName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrViewName);
			object returnItem = Invoker.MethodReturn(this, "RetrieveViewInfo", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="handlerType">Int32 handlerType</param>
		/// <param name="bstrCtrl">string bstrCtrl</param>
		/// <param name="bstrProperty">string bstrProperty</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public string GetEmbeddedObject(Int32 handlerType, string bstrCtrl, string bstrProperty)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(handlerType, bstrCtrl, bstrProperty);
			object returnItem = Invoker.MethodReturn(this, "GetEmbeddedObject", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="handlerType">Int32 handlerType</param>
		/// <param name="bstrCtrl">string bstrCtrl</param>
		/// <param name="bstrProperty">string bstrProperty</param>
		/// <param name="bstrExpression">string bstrExpression</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void SaveEmbeddedObject(Int32 handlerType, string bstrCtrl, string bstrProperty, string bstrExpression)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(handlerType, bstrCtrl, bstrProperty, bstrExpression);
			Invoker.Method(this, "SaveEmbeddedObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="handlerType">Int32 handlerType</param>
		/// <param name="bstrCtrl">string bstrCtrl</param>
		/// <param name="bstrProperty">string bstrProperty</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void DeleteEmbeddedObject(Int32 handlerType, string bstrCtrl, string bstrProperty)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(handlerType, bstrCtrl, bstrProperty);
			Invoker.Method(this, "DeleteEmbeddedObject", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="bstrFormName">string bstrFormName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public object GetFormBodyAndCss(string bstrFormName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrFormName);
			object returnItem = Invoker.MethodReturn(this, "GetFormBodyAndCss", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="accessObjectType">Int32 accessObjectType</param>
		/// <param name="varName">object varName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public bool IsValidAccessObjectName(Int32 accessObjectType, object varName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(accessObjectType, varName);
			object returnItem = Invoker.MethodReturn(this, "IsValidAccessObjectName", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}