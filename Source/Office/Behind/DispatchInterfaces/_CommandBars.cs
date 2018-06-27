using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface _CommandBars 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    public class _CommandBars : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi._CommandBars
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
                    _contractType = typeof(NetOffice.OfficeApi._CommandBars);
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
                    _type = typeof(_CommandBars);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _CommandBars() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862425.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [BaseResult]
        public virtual NetOffice.OfficeApi.CommandBarControl ActionControl
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OfficeApi.CommandBarControl>(this, "ActionControl");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863075.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBar ActiveMenuBar
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBar>(this, "ActiveMenuBar", typeof(NetOffice.OfficeApi.CommandBar));
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860520.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Count
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863160.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool DisplayTooltips
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayTooltips");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayTooltips", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864956.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool DisplayKeysInTooltips
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayKeysInTooltips");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayKeysInTooltips", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public virtual NetOffice.OfficeApi.CommandBar this[object index]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBar>(this, "Item", typeof(NetOffice.OfficeApi.CommandBar), index);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864068.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool LargeButtons
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "LargeButtons");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LargeButtons", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864076.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoMenuAnimation MenuAnimationStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoMenuAnimation>(this, "MenuAnimationStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "MenuAnimationStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862543.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="ids">Int32 ids</param>
        /// <param name="pbstrName">string pbstrName</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 get_IdsString(Int32 ids, out string pbstrName)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true);
            pbstrName = string.Empty;
            object[] paramsArray = new object[] { ids, pbstrName };

            Int32 returnItem = InvokerService.InvokeInternal.ExecuteInt32PropertyGetExtended(this, "IdsString", paramsArray, modifiers);
            
            pbstrName = paramsArray[1] as string;
            return returnItem;
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_IdsString
        /// </summary>
        /// <param name="ids">Int32 ids</param>
        /// <param name="pbstrName">string pbstrName</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_IdsString")]
        public virtual Int32 IdsString(Int32 ids, out string pbstrName)
        {
            return get_IdsString(ids, out pbstrName);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="tmc">Int32 tmc</param>
        /// <param name="pbstrName">string pbstrName</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 get_TmcGetName(Int32 tmc, out string pbstrName)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true);
            pbstrName = string.Empty;
            object[] paramsArray = new object[] { tmc, pbstrName };

            Int32 returnItem = InvokerService.InvokeInternal.ExecuteInt32PropertyGetExtended(this, "TmcGetName", paramsArray, modifiers);

            pbstrName = paramsArray[1] as string;
            return returnItem;
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_TmcGetName
        /// </summary>
        /// <param name="tmc">Int32 tmc</param>
        /// <param name="pbstrName">string pbstrName</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), Redirect("get_TmcGetName")]
        public virtual Int32 TmcGetName(Int32 tmc, out string pbstrName)
        {
            return get_TmcGetName(tmc, out pbstrName);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860590.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool AdaptiveMenus
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AdaptiveMenus");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AdaptiveMenus", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860823.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool DisplayFonts
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayFonts");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayFonts", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864631.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual bool DisableCustomize
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisableCustomize");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisableCustomize", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863405.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual bool DisableAskAQuestionDropdown
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisableAskAQuestionDropdown");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisableAskAQuestionDropdown", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861773.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="position">optional object position</param>
        /// <param name="menuBar">optional object menuBar</param>
        /// <param name="temporary">optional object temporary</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBar Add(object name, object position, object menuBar, object temporary)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "Add", typeof(NetOffice.OfficeApi.CommandBar), name, position, menuBar, temporary);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861773.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBar Add()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "Add", typeof(NetOffice.OfficeApi.CommandBar));
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861773.aspx </remarks>
        /// <param name="name">optional object name</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBar Add(object name)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "Add", typeof(NetOffice.OfficeApi.CommandBar), name);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861773.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="position">optional object position</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBar Add(object name, object position)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "Add", typeof(NetOffice.OfficeApi.CommandBar), name, position);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861773.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="position">optional object position</param>
        /// <param name="menuBar">optional object menuBar</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBar Add(object name, object position, object menuBar)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "Add", typeof(NetOffice.OfficeApi.CommandBar), name, position, menuBar);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860267.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="id">optional object id</param>
        /// <param name="tag">optional object tag</param>
        /// <param name="visible">optional object visible</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [BaseResult]
        public virtual NetOffice.OfficeApi.CommandBarControl FindControl(object type, object id, object tag, object visible)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "FindControl", type, id, tag, visible);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860267.aspx </remarks>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBarControl FindControl()
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "FindControl");
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860267.aspx </remarks>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBarControl FindControl(object type)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "FindControl", type);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860267.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="id">optional object id</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBarControl FindControl(object type, object id)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "FindControl", type, id);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860267.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="id">optional object id</param>
        /// <param name="tag">optional object tag</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBarControl FindControl(object type, object id, object tag)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "FindControl", type, id, tag);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861062.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ReleaseFocus()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ReleaseFocus");
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862556.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="id">optional object id</param>
        /// <param name="tag">optional object tag</param>
        /// <param name="visible">optional object visible</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBarControls FindControls(object type, object id, object tag, object visible)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControls>(this, "FindControls", typeof(NetOffice.OfficeApi.CommandBarControls), type, id, tag, visible);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862556.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBarControls FindControls()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControls>(this, "FindControls", typeof(NetOffice.OfficeApi.CommandBarControls));
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862556.aspx </remarks>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBarControls FindControls(object type)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControls>(this, "FindControls", typeof(NetOffice.OfficeApi.CommandBarControls), type);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862556.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="id">optional object id</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBarControls FindControls(object type, object id)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControls>(this, "FindControls", typeof(NetOffice.OfficeApi.CommandBarControls), type, id);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862556.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="id">optional object id</param>
        /// <param name="tag">optional object tag</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBarControls FindControls(object type, object id, object tag)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControls>(this, "FindControls", typeof(NetOffice.OfficeApi.CommandBarControls), type, id, tag);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="tbidOrName">optional object tbidOrName</param>
        /// <param name="position">optional object position</param>
        /// <param name="menuBar">optional object menuBar</param>
        /// <param name="temporary">optional object temporary</param>
        /// <param name="tbtrProtection">optional object tbtrProtection</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName, object position, object menuBar, object temporary, object tbtrProtection)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "AddEx", typeof(NetOffice.OfficeApi.CommandBar), new object[] { tbidOrName, position, menuBar, temporary, tbtrProtection });
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBar AddEx()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "AddEx", typeof(NetOffice.OfficeApi.CommandBar));
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="tbidOrName">optional object tbidOrName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "AddEx", typeof(NetOffice.OfficeApi.CommandBar), tbidOrName);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="tbidOrName">optional object tbidOrName</param>
        /// <param name="position">optional object position</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName, object position)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "AddEx", typeof(NetOffice.OfficeApi.CommandBar), tbidOrName, position);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="tbidOrName">optional object tbidOrName</param>
        /// <param name="position">optional object position</param>
        /// <param name="menuBar">optional object menuBar</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName, object position, object menuBar)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "AddEx", typeof(NetOffice.OfficeApi.CommandBar), tbidOrName, position, menuBar);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="tbidOrName">optional object tbidOrName</param>
        /// <param name="position">optional object position</param>
        /// <param name="menuBar">optional object menuBar</param>
        /// <param name="temporary">optional object temporary</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName, object position, object menuBar, object temporary)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "AddEx", typeof(NetOffice.OfficeApi.CommandBar), tbidOrName, position, menuBar, temporary);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862419.aspx </remarks>
        /// <param name="idMso">string idMso</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ExecuteMso(string idMso)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExecuteMso", idMso);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862202.aspx </remarks>
        /// <param name="idMso">string idMso</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool GetEnabledMso(string idMso)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "GetEnabledMso", idMso);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863712.aspx </remarks>
        /// <param name="idMso">string idMso</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool GetVisibleMso(string idMso)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "GetVisibleMso", idMso);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863149.aspx </remarks>
        /// <param name="idMso">string idMso</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool GetPressedMso(string idMso)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "GetPressedMso", idMso);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860585.aspx </remarks>
        /// <param name="idMso">string idMso</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string GetLabelMso(string idMso)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetLabelMso", idMso);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860790.aspx </remarks>
        /// <param name="idMso">string idMso</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string GetScreentipMso(string idMso)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetScreentipMso", idMso);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864975.aspx </remarks>
        /// <param name="idMso">string idMso</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string GetSupertipMso(string idMso)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetSupertipMso", idMso);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861156.aspx </remarks>
        /// <param name="idMso">string idMso</param>
        /// <param name="width">Int32 width</param>
        /// <param name="height">Int32 height</param>
        [SupportByVersion("Office", 12, 14, 15, 16), NativeResult]
        public virtual stdole.Picture GetImageMso(string idMso, Int32 width, Int32 height)
        {
            object[] paramsArray = new object[] { idMso, width, height };
            object returnItem = InvokerService.InvokeInternal.ExecuteObjectMethodGet(this, "GetImageMso", paramsArray);
            return returnItem as stdole.Picture;
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863478.aspx </remarks>
        /// <param name="hwnd">Int32 hwnd</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual void CommitRenderingTransaction(Int32 hwnd)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CommitRenderingTransaction", hwnd);
        }

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.CommandBar>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.CommandBar>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.CommandBar>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.CommandBar>

        /// <summary>
        /// SupportByVersion Office, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.OfficeApi.CommandBar> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.CommandBar item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// SupportByVersion Office, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}
