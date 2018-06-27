using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.AccessApi;

namespace NetOffice.AccessApi.Behind
{
	/// <summary>
	/// DispatchInterface _DesignerReceiver 
	/// SupportByVersion Access, 15, 16
	/// </summary>
	[SupportByVersion("Access", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class _DesignerReceiver : COMObject, NetOffice.AccessApi._DesignerReceiver
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
                    _contractType = typeof(NetOffice.AccessApi._DesignerReceiver);
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
                    _type = typeof(_DesignerReceiver);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _DesignerReceiver() : base()
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
		public virtual void SetReady()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetReady");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void RefreshRibbon()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RefreshRibbon");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="builder">Int16 builder</param>
		/// <param name="bstrBuilderValue">string bstrBuilderValue</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual object LaunchBuilder(Int16 builder, string bstrBuilderValue)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LaunchBuilder", builder, bstrBuilderValue);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="propertytype">Int16 propertytype</param>
		/// <param name="inputData">optional object inputData</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual object RetrievePropertyValues(Int16 propertytype, object inputData)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "RetrievePropertyValues", propertytype, inputData);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="propertytype">Int16 propertytype</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 15, 16)]
		public virtual object RetrievePropertyValues(Int16 propertytype)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "RetrievePropertyValues", propertytype);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void RecordSourceUpdated(string bstrRecordSource)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "RecordSourceUpdated", bstrRecordSource);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual string GetEmbeddedRecordSourceSQL()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetEmbeddedRecordSourceSQL");
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual string SetEmbeddedRecordSourceSQL(string bstrRecordSource)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "SetEmbeddedRecordSourceSQL", bstrRecordSource);
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
		public virtual void QuickCreateView(string bstrViewName, Int16 viewtype, string bstrRecordSource, bool fStandalone)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "QuickCreateView", bstrViewName, viewtype, bstrRecordSource, fStandalone);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrExpression">string bstrExpression</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual object IsExpressionValid(string bstrExpression)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "IsExpressionValid", bstrExpression);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrOldCtrlID">string bstrOldCtrlID</param>
		/// <param name="bstrNewCtrlID">string bstrNewCtrlID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void NotifyControlIDChanged(string bstrOldCtrlID, string bstrNewCtrlID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NotifyControlIDChanged", bstrOldCtrlID, bstrNewCtrlID);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual object RetrieveControlSourcesInfo(string bstrRecordSource)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "RetrieveControlSourcesInfo", bstrRecordSource);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="accessObjectType">Int32 accessObjectType</param>
		/// <param name="accessObjectName">string accessObjectName</param>
		/// <param name="fDesignMode">bool fDesignMode</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void OpenAccessObject(Int32 accessObjectType, string accessObjectName, bool fDesignMode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenAccessObject", accessObjectType, accessObjectName, fDesignMode);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="accessObjectType">Int32 accessObjectType</param>
		/// <param name="accessObjectName">string accessObjectName</param>
		/// <param name="pdispDependentObjectTypeNamePairArray">object pdispDependentObjectTypeNamePairArray</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual bool DeleteAccessObject(Int32 accessObjectType, string accessObjectName, object pdispDependentObjectTypeNamePairArray)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "DeleteAccessObject", accessObjectType, accessObjectName, pdispDependentObjectTypeNamePairArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="accessObjectType">Int32 accessObjectType</param>
		/// <param name="accessObjectName">string accessObjectName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void DuplicateAccessObject(Int32 accessObjectType, string accessObjectName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DuplicateAccessObject", accessObjectType, accessObjectName);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual object RetrieveRecordSourceInfo(string bstrRecordSource)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "RetrieveRecordSourceInfo", bstrRecordSource);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrViewName">string bstrViewName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual object RetrieveViewInfo(string bstrViewName)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "RetrieveViewInfo", bstrViewName);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="handlerType">Int32 handlerType</param>
		/// <param name="bstrCtrl">string bstrCtrl</param>
		/// <param name="bstrProperty">string bstrProperty</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual string GetEmbeddedObject(Int32 handlerType, string bstrCtrl, string bstrProperty)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetEmbeddedObject", handlerType, bstrCtrl, bstrProperty);
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
		public virtual void SaveEmbeddedObject(Int32 handlerType, string bstrCtrl, string bstrProperty, string bstrExpression)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveEmbeddedObject", handlerType, bstrCtrl, bstrProperty, bstrExpression);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="handlerType">Int32 handlerType</param>
		/// <param name="bstrCtrl">string bstrCtrl</param>
		/// <param name="bstrProperty">string bstrProperty</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void DeleteEmbeddedObject(Int32 handlerType, string bstrCtrl, string bstrProperty)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteEmbeddedObject", handlerType, bstrCtrl, bstrProperty);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrFormName">string bstrFormName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual object GetFormBodyAndCss(string bstrFormName)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetFormBodyAndCss", bstrFormName);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="accessObjectType">Int32 accessObjectType</param>
		/// <param name="varName">object varName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual bool IsValidAccessObjectName(Int32 accessObjectType, object varName)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsValidAccessObjectName", accessObjectType, varName);
		}

		#endregion

		#pragma warning restore
	}
}

