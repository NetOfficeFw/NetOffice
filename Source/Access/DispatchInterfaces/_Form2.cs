using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _Form2 
	/// SupportByVersion Access, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method)]
	public class _Form2 : _Form, IEnumerableProvider<object>
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
                    _type = typeof(_Form2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Form2(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Form2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Form2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Form2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Form2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Form2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Form2() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Form2(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public byte DatasheetBorderLineStyle
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "DatasheetBorderLineStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetBorderLineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public byte DatasheetColumnHeaderUnderlineStyle
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "DatasheetColumnHeaderUnderlineStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DatasheetColumnHeaderUnderlineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public byte HorizontalDatasheetGridlineStyle
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "HorizontalDatasheetGridlineStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HorizontalDatasheetGridlineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public byte VerticalDatasheetGridlineStyle
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "VerticalDatasheetGridlineStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "VerticalDatasheetGridlineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int16 WindowTop
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "WindowTop");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public Int16 WindowLeft
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "WindowLeft");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string OnUndo
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnUndo");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnUndo", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnRecordExit
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnRecordExit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnRecordExit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public NetOffice.OWC10Api.PivotTable PivotTable
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PivotTable>(this, "PivotTable", NetOffice.OWC10Api.PivotTable.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public NetOffice.OWC10Api.ChartSpace ChartSpace
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChartSpace>(this, "ChartSpace", NetOffice.OWC10Api.ChartSpace.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.AccessApi._Printer Printer
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.AccessApi._Printer>(this, "Printer");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Printer", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool Moveable
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Moveable");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Moveable", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BeginBatchEdit
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeginBatchEdit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeginBatchEdit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string UndoBatchEdit
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UndoBatchEdit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UndoBatchEdit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BeforeBeginTransaction
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeBeginTransaction");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeBeginTransaction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string AfterBeginTransaction
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterBeginTransaction");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterBeginTransaction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string BeforeCommitTransaction
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeCommitTransaction");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeCommitTransaction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string AfterCommitTransaction
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterCommitTransaction");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterCommitTransaction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string RollbackTransaction
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RollbackTransaction");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RollbackTransaction", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool AllowFormView
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowFormView");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowFormView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool AllowDatasheetView
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowDatasheetView");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowDatasheetView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool AllowPivotTableView
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowPivotTableView");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowPivotTableView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool AllowPivotChartView
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowPivotChartView");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowPivotChartView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string OnConnect
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnConnect");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnConnect", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string OnDisconnect
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDisconnect");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDisconnect", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string PivotTableChange
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PivotTableChange");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PivotTableChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string Query
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Query");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Query", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string BeforeQuery
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeQuery");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeQuery", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string SelectionChange
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SelectionChange");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SelectionChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string CommandBeforeExecute
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CommandBeforeExecute");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandBeforeExecute", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string CommandChecked
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CommandChecked");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandChecked", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string CommandEnabled
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CommandEnabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandEnabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string CommandExecute
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CommandExecute");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandExecute", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string DataSetChange
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataSetChange");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataSetChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string BeforeScreenTip
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeScreenTip");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeScreenTip", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string AfterFinalRender
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterFinalRender");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterFinalRender", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string AfterRender
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterRender");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterRender", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string AfterLayout
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AfterLayout");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AfterLayout", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string BeforeRender
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BeforeRender");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BeforeRender", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string MouseWheel
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MouseWheel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MouseWheel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string ViewChange
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ViewChange");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ViewChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string DataChange
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataChange");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool FetchDefaults
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FetchDefaults");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FetchDefaults", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool BatchUpdates
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BatchUpdates");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BatchUpdates", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public byte CommitOnClose
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "CommitOnClose");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommitOnClose", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool CommitOnNavigation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CommitOnNavigation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommitOnNavigation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool UseDefaultPrinter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseDefaultPrinter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseDefaultPrinter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public string RecordSourceQualifier
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RecordSourceQualifier");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordSourceQualifier", value);
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
		public void Move(object left, object top, object width, object height)
		{
			 Factory.ExecuteMethod(this, "Move", left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public void Move(object left)
		{
			 Factory.ExecuteMethod(this, "Move", left);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public void Move(object left, object top)
		{
			 Factory.ExecuteMethod(this, "Move", left, top);
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public void Move(object left, object top, object width)
		{
			 Factory.ExecuteMethod(this, "Move", left, top, width);
		}

        #endregion
      
        #region IEnumerableProvider<object>

        ICOMObject IEnumerableProvider<object>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this);
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
        public IEnumerator<object> GetEnumerator()
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
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion

		#pragma warning restore
	}
}
