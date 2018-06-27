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
	/// DispatchInterface _Report3 
	/// SupportByVersion Access, 12,14,15,16
	/// </summary>
	public class _Report3 : _Report2, NetOffice.AccessApi._Report3
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
                    _contractType = typeof(NetOffice.AccessApi._Report3);
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
                    _type = typeof(_Report3);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _Report3() : base()
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
		public virtual byte DefaultView
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "DefaultView");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual bool AllowReportView
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowReportView");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowReportView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual byte ScrollBars
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "ScrollBars");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScrollBars", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual byte Cycle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "Cycle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Cycle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool AllowDesignChanges
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowDesignChanges");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowDesignChanges", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string OnCurrent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnCurrent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnCurrent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual bool KeyPreview
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "KeyPreview");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "KeyPreview", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual Int32 TimerInterval
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "TimerInterval");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TimerInterval", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual Int16 CurrentView
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "CurrentView");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CurrentView", value);
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
		public virtual string OnNoDataMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnNoDataMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnNoDataMacro", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnPageMacro
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnPageMacro");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnPageMacro", value);
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
		public virtual bool ShowPageMargins
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowPageMargins");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowPageMargins", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual bool FitToPage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FitToPage");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FitToPage", value);
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
		public virtual string OnLoad
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnLoad");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnLoad", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string OnResize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnResize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnResize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string OnUnload
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnUnload");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnUnload", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string OnGotFocus
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnGotFocus");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnGotFocus", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string OnLostFocus
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnLostFocus");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnLostFocus", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string OnClick
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnClick");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnClick", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string OnDblClick
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnDblClick");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnDblClick", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string OnMouseDown
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnMouseDown");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnMouseDown", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string OnMouseMove
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnMouseMove");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnMouseMove", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string OnMouseUp
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnMouseUp");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnMouseUp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string OnKeyDown
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnKeyDown");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnKeyDown", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string OnKeyUp
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnKeyUp");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnKeyUp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string OnKeyPress
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnKeyPress");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnKeyPress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string OnFilter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnFilter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string OnApplyFilter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnApplyFilter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnApplyFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual string OnTimer
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnTimer");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnTimer", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
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

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual void Requery()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Requery");
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

