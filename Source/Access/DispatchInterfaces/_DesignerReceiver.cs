using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _DesignerReceiver 
	/// SupportByVersion Access, 15, 16
	/// </summary>
	[SupportByVersion("Access", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class _DesignerReceiver : COMObject
	{
		#pragma warning disable

		#region Type Information

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
                    _type = typeof(_DesignerReceiver);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _DesignerReceiver(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void SetReady()
		{
			 Factory.ExecuteMethod(this, "SetReady");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void RefreshRibbon()
		{
			 Factory.ExecuteMethod(this, "RefreshRibbon");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="builder">Int16 builder</param>
		/// <param name="bstrBuilderValue">string bstrBuilderValue</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public object LaunchBuilder(Int16 builder, string bstrBuilderValue)
		{
			return Factory.ExecuteVariantMethodGet(this, "LaunchBuilder", builder, bstrBuilderValue);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="propertytype">Int16 propertytype</param>
		/// <param name="inputData">optional object inputData</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public object RetrievePropertyValues(Int16 propertytype, object inputData)
		{
			return Factory.ExecuteVariantMethodGet(this, "RetrievePropertyValues", propertytype, inputData);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="propertytype">Int16 propertytype</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 15, 16)]
		public object RetrievePropertyValues(Int16 propertytype)
		{
			return Factory.ExecuteVariantMethodGet(this, "RetrievePropertyValues", propertytype);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void RecordSourceUpdated(string bstrRecordSource)
		{
			 Factory.ExecuteMethod(this, "RecordSourceUpdated", bstrRecordSource);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public string GetEmbeddedRecordSourceSQL()
		{
			return Factory.ExecuteStringMethodGet(this, "GetEmbeddedRecordSourceSQL");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public string SetEmbeddedRecordSourceSQL(string bstrRecordSource)
		{
			return Factory.ExecuteStringMethodGet(this, "SetEmbeddedRecordSourceSQL", bstrRecordSource);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrViewName">string bstrViewName</param>
		/// <param name="viewtype">Int16 viewtype</param>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		/// <param name="fStandalone">bool fStandalone</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void QuickCreateView(string bstrViewName, Int16 viewtype, string bstrRecordSource, bool fStandalone)
		{
			 Factory.ExecuteMethod(this, "QuickCreateView", bstrViewName, viewtype, bstrRecordSource, fStandalone);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrExpression">string bstrExpression</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public object IsExpressionValid(string bstrExpression)
		{
			return Factory.ExecuteVariantMethodGet(this, "IsExpressionValid", bstrExpression);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrOldCtrlID">string bstrOldCtrlID</param>
		/// <param name="bstrNewCtrlID">string bstrNewCtrlID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void NotifyControlIDChanged(string bstrOldCtrlID, string bstrNewCtrlID)
		{
			 Factory.ExecuteMethod(this, "NotifyControlIDChanged", bstrOldCtrlID, bstrNewCtrlID);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public object RetrieveControlSourcesInfo(string bstrRecordSource)
		{
			return Factory.ExecuteVariantMethodGet(this, "RetrieveControlSourcesInfo", bstrRecordSource);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="accessObjectType">Int32 accessObjectType</param>
		/// <param name="accessObjectName">string accessObjectName</param>
		/// <param name="fDesignMode">bool fDesignMode</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void OpenAccessObject(Int32 accessObjectType, string accessObjectName, bool fDesignMode)
		{
			 Factory.ExecuteMethod(this, "OpenAccessObject", accessObjectType, accessObjectName, fDesignMode);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="accessObjectType">Int32 accessObjectType</param>
		/// <param name="accessObjectName">string accessObjectName</param>
		/// <param name="pdispDependentObjectTypeNamePairArray">object pdispDependentObjectTypeNamePairArray</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public bool DeleteAccessObject(Int32 accessObjectType, string accessObjectName, object pdispDependentObjectTypeNamePairArray)
		{
			return Factory.ExecuteBoolMethodGet(this, "DeleteAccessObject", accessObjectType, accessObjectName, pdispDependentObjectTypeNamePairArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="accessObjectType">Int32 accessObjectType</param>
		/// <param name="accessObjectName">string accessObjectName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void DuplicateAccessObject(Int32 accessObjectType, string accessObjectName)
		{
			 Factory.ExecuteMethod(this, "DuplicateAccessObject", accessObjectType, accessObjectName);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public object RetrieveRecordSourceInfo(string bstrRecordSource)
		{
			return Factory.ExecuteVariantMethodGet(this, "RetrieveRecordSourceInfo", bstrRecordSource);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrViewName">string bstrViewName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public object RetrieveViewInfo(string bstrViewName)
		{
			return Factory.ExecuteVariantMethodGet(this, "RetrieveViewInfo", bstrViewName);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="handlerType">Int32 handlerType</param>
		/// <param name="bstrCtrl">string bstrCtrl</param>
		/// <param name="bstrProperty">string bstrProperty</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public string GetEmbeddedObject(Int32 handlerType, string bstrCtrl, string bstrProperty)
		{
			return Factory.ExecuteStringMethodGet(this, "GetEmbeddedObject", handlerType, bstrCtrl, bstrProperty);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="handlerType">Int32 handlerType</param>
		/// <param name="bstrCtrl">string bstrCtrl</param>
		/// <param name="bstrProperty">string bstrProperty</param>
		/// <param name="bstrExpression">string bstrExpression</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void SaveEmbeddedObject(Int32 handlerType, string bstrCtrl, string bstrProperty, string bstrExpression)
		{
			 Factory.ExecuteMethod(this, "SaveEmbeddedObject", handlerType, bstrCtrl, bstrProperty, bstrExpression);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="handlerType">Int32 handlerType</param>
		/// <param name="bstrCtrl">string bstrCtrl</param>
		/// <param name="bstrProperty">string bstrProperty</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void DeleteEmbeddedObject(Int32 handlerType, string bstrCtrl, string bstrProperty)
		{
			 Factory.ExecuteMethod(this, "DeleteEmbeddedObject", handlerType, bstrCtrl, bstrProperty);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrFormName">string bstrFormName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public object GetFormBodyAndCss(string bstrFormName)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetFormBodyAndCss", bstrFormName);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="accessObjectType">Int32 accessObjectType</param>
		/// <param name="varName">object varName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public bool IsValidAccessObjectName(Int32 accessObjectType, object varName)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsValidAccessObjectName", accessObjectType, varName);
		}

		#endregion

		#pragma warning restore
	}
}
