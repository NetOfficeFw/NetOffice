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
	/// DispatchInterface _Form3 
	/// SupportByVersion Access, 12,14,15,16
	/// </summary>
	public class _Form3 : _Form2, NetOffice.AccessApi._Form3
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
                    _contractType = typeof(NetOffice.AccessApi._Form3);
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
                    _type = typeof(_Form3);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _Form3() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 12
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 12)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.AccessApi.Section get__SectionOld(object index)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Section>(this, "_SectionOld", typeof(NetOffice.AccessApi.Section), index);
		}

		/// <summary>
		/// SupportByVersion Access 12
		/// Alias for get__SectionOld
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 12), Redirect("get__SectionOld")]
		public virtual NetOffice.AccessApi.Section _SectionOld(object index)
		{
			return get__SectionOld(index);
		}
        
		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16), ProxyResult]
		public virtual object PivotTable
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "PivotTable");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Access", 10,11,12,14,15,16), ProxyResult]
		public virtual object ChartSpace
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ChartSpace");
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual bool FilterOnLoad
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FilterOnLoad");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FilterOnLoad", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual bool OrderByOnLoad
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "OrderByOnLoad");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OrderByOnLoad", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Enums.AcSplitFormOrientation SplitFormOrientation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcSplitFormOrientation>(this, "SplitFormOrientation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SplitFormOrientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Enums.AcSplitFormDatasheet SplitFormDatasheet
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcSplitFormDatasheet>(this, "SplitFormDatasheet");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SplitFormDatasheet", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual bool SplitFormSplitterBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SplitFormSplitterBar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SplitFormSplitterBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.Enums.AcSplitFormPrinting SplitFormPrinting
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcSplitFormPrinting>(this, "SplitFormPrinting");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SplitFormPrinting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual bool SplitFormSplitterBarSave
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SplitFormSplitterBarSave");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SplitFormSplitterBarSave", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string NavigationCaption
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "NavigationCaption");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NavigationCaption", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnCurrentMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnCurrentMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnCurrentMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string BeforeInsertMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeforeInsertMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeforeInsertMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string AfterInsertMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AfterInsertMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AfterInsertMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string BeforeUpdateMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeforeUpdateMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeforeUpdateMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string AfterUpdateMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AfterUpdateMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AfterUpdateMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnDirtyMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnDirtyMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnDirtyMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnDeleteMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnDeleteMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnDeleteMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string BeforeDelConfirmMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeforeDelConfirmMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeforeDelConfirmMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string AfterDelConfirmMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AfterDelConfirmMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AfterDelConfirmMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnOpenMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnOpenMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnOpenMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnLoadMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnLoadMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnLoadMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnResizeMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnResizeMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnResizeMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnUnloadMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnUnloadMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnUnloadMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnCloseMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnCloseMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnCloseMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnActivateMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnActivateMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnActivateMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnDeactivateMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnDeactivateMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnDeactivateMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnGotFocusMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnGotFocusMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnGotFocusMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnLostFocusMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnLostFocusMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnLostFocusMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnClickMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnClickMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnClickMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnDblClickMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnDblClickMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnDblClickMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnMouseDownMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnMouseDownMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnMouseDownMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnMouseMoveMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnMouseMoveMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnMouseMoveMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnMouseUpMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnMouseUpMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnMouseUpMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnKeyDownMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnKeyDownMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnKeyDownMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnKeyUpMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnKeyUpMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnKeyUpMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnKeyPressMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnKeyPressMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnKeyPressMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnErrorMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnErrorMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnErrorMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnFilterMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnFilterMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnFilterMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnApplyFilterMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnApplyFilterMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnApplyFilterMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnTimerMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnTimerMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnTimerMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnUndoMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnUndoMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnUndoMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnRecordExitMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnRecordExitMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnRecordExitMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string BeginBatchEditMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeginBatchEditMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeginBatchEditMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string UndoBatchEditMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "UndoBatchEditMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UndoBatchEditMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string BeforeBeginTransactionMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeforeBeginTransactionMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeforeBeginTransactionMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string AfterBeginTransactionMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AfterBeginTransactionMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AfterBeginTransactionMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string BeforeCommitTransactionMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeforeCommitTransactionMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeforeCommitTransactionMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string AfterCommitTransactionMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AfterCommitTransactionMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AfterCommitTransactionMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string RollbackTransactionMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RollbackTransactionMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RollbackTransactionMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnConnectMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnConnectMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnConnectMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnDisconnectMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnDisconnectMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnDisconnectMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string PivotTableChangeMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PivotTableChangeMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PivotTableChangeMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string QueryMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "QueryMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "QueryMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string BeforeQueryMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeforeQueryMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeforeQueryMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string SelectionChangeMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SelectionChangeMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SelectionChangeMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string CommandBeforeExecuteMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CommandBeforeExecuteMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CommandBeforeExecuteMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string CommandCheckedMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CommandCheckedMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CommandCheckedMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string CommandEnabledMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CommandEnabledMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CommandEnabledMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string CommandExecuteMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CommandExecuteMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CommandExecuteMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string DataSetChangeMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DataSetChangeMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DataSetChangeMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string BeforeScreenTipMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeforeScreenTipMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeforeScreenTipMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string AfterFinalRenderMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AfterFinalRenderMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AfterFinalRenderMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string AfterRenderMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AfterRenderMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AfterRenderMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string AfterLayoutMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AfterLayoutMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AfterLayoutMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string BeforeRenderMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeforeRenderMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeforeRenderMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string MouseWheelMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "MouseWheelMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MouseWheelMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string ViewChangeMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ViewChangeMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ViewChangeMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string DataChangeMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DataChangeMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DataChangeMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual bool AllowLayoutView
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowLayoutView");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowLayoutView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual Int32 DatasheetAlternateBackColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DatasheetAlternateBackColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DatasheetAlternateBackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual byte DisplayOnSharePointSite
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "DisplayOnSharePointSite");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayOnSharePointSite", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual Int32 SplitFormSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SplitFormSize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SplitFormSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.AccessApi._Section get_Section(object index)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi._Section>(this, "Section", typeof(NetOffice.AccessApi._Section), index);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Section
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_Section")]
		public virtual NetOffice.AccessApi._Section Section(object index)
		{
			return get_Section(index);
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string RibbonName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RibbonName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RibbonName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual bool FitToScreen
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FitToScreen");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FitToScreen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.AccessApi.Section get_SectionOld(object index)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Section>(this, "SectionOld", typeof(NetOffice.AccessApi.Section), index);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Alias for get_SectionOld
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 14,15,16), Redirect("get_SectionOld")]
		public virtual NetOffice.AccessApi.Section SectionOld(object index)
		{
			return get_SectionOld(index);
		}

        #endregion

        #region Methods

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
        /// SupportByVersion Access, 12,14,15,16
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual IEnumerator<object> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (object item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Access, 12,14,15,16
        /// </summary>
        [SupportByVersion("Access", 12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this, true);
		}

		#endregion

		#pragma warning restore
	}
}

