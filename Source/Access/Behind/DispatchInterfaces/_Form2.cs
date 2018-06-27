using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.AccessApi;

namespace NetOffice.AccessApi.Behind
{
	/// <summary>
	/// DispatchInterface _Form2 
	/// SupportByVersion Access, 10,11,12,14,15,16
	/// </summary>
	public class _Form2 : _Form, NetOffice.AccessApi._Form2
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
                    _contractType = typeof(NetOffice.AccessApi._Form2);
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
                    _type = typeof(_Form2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _Form2() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual byte DatasheetBorderLineStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "DatasheetBorderLineStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DatasheetBorderLineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual byte DatasheetColumnHeaderUnderlineStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "DatasheetColumnHeaderUnderlineStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DatasheetColumnHeaderUnderlineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual byte HorizontalDatasheetGridlineStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "HorizontalDatasheetGridlineStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HorizontalDatasheetGridlineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual byte VerticalDatasheetGridlineStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "VerticalDatasheetGridlineStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "VerticalDatasheetGridlineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual Int16 WindowTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "WindowTop");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual Int16 WindowLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "WindowLeft");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string OnUndo
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnUndo");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnUndo", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnRecordExit
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnRecordExit");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnRecordExit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual NetOffice.OWC10Api.PivotTable PivotTable
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotTable>(this, "PivotTable", typeof(NetOffice.OWC10Api.PivotTable));
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual NetOffice.OWC10Api.ChartSpace ChartSpace
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChartSpace>(this, "ChartSpace", typeof(NetOffice.OWC10Api.ChartSpace));
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.AccessApi._Printer Printer
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.AccessApi._Printer>(this, "Printer");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Printer", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual bool Moveable
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Moveable");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Moveable", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string BeginBatchEdit
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeginBatchEdit");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeginBatchEdit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string UndoBatchEdit
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "UndoBatchEdit");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UndoBatchEdit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string BeforeBeginTransaction
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeforeBeginTransaction");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeforeBeginTransaction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string AfterBeginTransaction
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AfterBeginTransaction");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AfterBeginTransaction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string BeforeCommitTransaction
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeforeCommitTransaction");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeforeCommitTransaction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string AfterCommitTransaction
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AfterCommitTransaction");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AfterCommitTransaction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string RollbackTransaction
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RollbackTransaction");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RollbackTransaction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual bool AllowFormView
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowFormView");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowFormView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual bool AllowDatasheetView
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowDatasheetView");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowDatasheetView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual bool AllowPivotTableView
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowPivotTableView");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowPivotTableView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual bool AllowPivotChartView
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowPivotChartView");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowPivotChartView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string OnConnect
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnConnect");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnConnect", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string OnDisconnect
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnDisconnect");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnDisconnect", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string PivotTableChange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PivotTableChange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PivotTableChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string Query
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Query");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Query", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string BeforeQuery
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeforeQuery");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeforeQuery", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string SelectionChange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SelectionChange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SelectionChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string CommandBeforeExecute
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CommandBeforeExecute");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CommandBeforeExecute", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string CommandChecked
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CommandChecked");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CommandChecked", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string CommandEnabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CommandEnabled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CommandEnabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string CommandExecute
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CommandExecute");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CommandExecute", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string DataSetChange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DataSetChange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DataSetChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string BeforeScreenTip
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeforeScreenTip");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeforeScreenTip", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string AfterFinalRender
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AfterFinalRender");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AfterFinalRender", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string AfterRender
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AfterRender");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AfterRender", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string AfterLayout
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AfterLayout");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AfterLayout", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string BeforeRender
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeforeRender");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeforeRender", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string MouseWheel
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "MouseWheel");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MouseWheel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string ViewChange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ViewChange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ViewChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string DataChange
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DataChange");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DataChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual bool FetchDefaults
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FetchDefaults");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FetchDefaults", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool BatchUpdates
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "BatchUpdates");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BatchUpdates", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual byte CommitOnClose
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "CommitOnClose");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CommitOnClose", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool CommitOnNavigation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CommitOnNavigation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CommitOnNavigation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual bool UseDefaultPrinter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UseDefaultPrinter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UseDefaultPrinter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual string RecordSourceQualifier
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RecordSourceQualifier");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RecordSourceQualifier", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void Move(object left, object top, object width, object height)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Move", left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void Move(object left)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Move", left);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void Move(object left, object top)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Move", left, top);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual void Move(object left, object top, object width)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Move", left, top, width);
		}

        #endregion
      
        #region IEnumerableProvider<object>

        ICOMObject IEnumerableProvider<object>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this, false);
        }

        IEnumerable IEnumerableProvider<object>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, true);
        }

        #endregion

        #region IEnumerable<object>

        /// <summary>
        /// SupportByVersion Access, 10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Access", 10, 11, 12, 14, 15, 16)]
        public virtual IEnumerator<object> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (object item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Access, 10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Access", 10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this, true);
		}

		#endregion

		#pragma warning restore
	}
}


