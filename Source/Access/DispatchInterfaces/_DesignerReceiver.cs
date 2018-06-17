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
	[TypeId("4D2A337B-259D-44A6-A5C6-81A629228CCF")]
	public interface _DesignerReceiver : ICOMObject
	{
		#region Methods

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void SetReady();

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void RefreshRibbon();

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="builder">Int16 builder</param>
		/// <param name="bstrBuilderValue">string bstrBuilderValue</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		object LaunchBuilder(Int16 builder, string bstrBuilderValue);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="propertytype">Int16 propertytype</param>
		/// <param name="inputData">optional object inputData</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		object RetrievePropertyValues(Int16 propertytype, object inputData);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="propertytype">Int16 propertytype</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 15, 16)]
		object RetrievePropertyValues(Int16 propertytype);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void RecordSourceUpdated(string bstrRecordSource);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		string GetEmbeddedRecordSourceSQL();

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		string SetEmbeddedRecordSourceSQL(string bstrRecordSource);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrViewName">string bstrViewName</param>
		/// <param name="viewtype">Int16 viewtype</param>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		/// <param name="fStandalone">bool fStandalone</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void QuickCreateView(string bstrViewName, Int16 viewtype, string bstrRecordSource, bool fStandalone);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrExpression">string bstrExpression</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		object IsExpressionValid(string bstrExpression);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrOldCtrlID">string bstrOldCtrlID</param>
		/// <param name="bstrNewCtrlID">string bstrNewCtrlID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void NotifyControlIDChanged(string bstrOldCtrlID, string bstrNewCtrlID);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		object RetrieveControlSourcesInfo(string bstrRecordSource);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="accessObjectType">Int32 accessObjectType</param>
		/// <param name="accessObjectName">string accessObjectName</param>
		/// <param name="fDesignMode">bool fDesignMode</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void OpenAccessObject(Int32 accessObjectType, string accessObjectName, bool fDesignMode);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="accessObjectType">Int32 accessObjectType</param>
		/// <param name="accessObjectName">string accessObjectName</param>
		/// <param name="pdispDependentObjectTypeNamePairArray">object pdispDependentObjectTypeNamePairArray</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		bool DeleteAccessObject(Int32 accessObjectType, string accessObjectName, object pdispDependentObjectTypeNamePairArray);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="accessObjectType">Int32 accessObjectType</param>
		/// <param name="accessObjectName">string accessObjectName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void DuplicateAccessObject(Int32 accessObjectType, string accessObjectName);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrRecordSource">string bstrRecordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		object RetrieveRecordSourceInfo(string bstrRecordSource);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrViewName">string bstrViewName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		object RetrieveViewInfo(string bstrViewName);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="handlerType">Int32 handlerType</param>
		/// <param name="bstrCtrl">string bstrCtrl</param>
		/// <param name="bstrProperty">string bstrProperty</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		string GetEmbeddedObject(Int32 handlerType, string bstrCtrl, string bstrProperty);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="handlerType">Int32 handlerType</param>
		/// <param name="bstrCtrl">string bstrCtrl</param>
		/// <param name="bstrProperty">string bstrProperty</param>
		/// <param name="bstrExpression">string bstrExpression</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void SaveEmbeddedObject(Int32 handlerType, string bstrCtrl, string bstrProperty, string bstrExpression);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="handlerType">Int32 handlerType</param>
		/// <param name="bstrCtrl">string bstrCtrl</param>
		/// <param name="bstrProperty">string bstrProperty</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		void DeleteEmbeddedObject(Int32 handlerType, string bstrCtrl, string bstrProperty);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="bstrFormName">string bstrFormName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		object GetFormBodyAndCss(string bstrFormName);

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="accessObjectType">Int32 accessObjectType</param>
		/// <param name="varName">object varName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		bool IsValidAccessObjectName(Int32 accessObjectType, object varName);

		#endregion
	}
}
