using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface Assistant 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class Assistant : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.Assistant
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
                    _contractType = typeof(NetOffice.OfficeApi.Assistant);
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
                    _type = typeof(Assistant);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Assistant() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
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
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Top
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Top");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Top", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Left
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Left");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Left", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Balloon NewBalloon
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Balloon>(this, "NewBalloon", typeof(NetOffice.OfficeApi.Balloon));
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoBalloonErrorType BalloonError
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoBalloonErrorType>(this, "BalloonError");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Visible
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Visible");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Visible", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoAnimationType Animation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoAnimationType>(this, "Animation");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Animation", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Reduced
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Reduced");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Reduced", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool AssistWithHelp
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AssistWithHelp");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AssistWithHelp", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool AssistWithWizards
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AssistWithWizards");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AssistWithWizards", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool AssistWithAlerts
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AssistWithAlerts");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AssistWithAlerts", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool MoveWhenInTheWay
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MoveWhenInTheWay");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MoveWhenInTheWay", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Sounds
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Sounds");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Sounds", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool FeatureTips
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FeatureTips");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FeatureTips", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool MouseTips
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MouseTips");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MouseTips", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool KeyboardShortcutTips
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "KeyboardShortcutTips");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "KeyboardShortcutTips", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool HighPriorityTips
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HighPriorityTips");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HighPriorityTips", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool TipOfDay
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "TipOfDay");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TipOfDay", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool GuessHelp
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "GuessHelp");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GuessHelp", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool SearchWhenProgramming
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SearchWhenProgramming");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SearchWhenProgramming", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string Item
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Item");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string FileName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FileName");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FileName", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string Name
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool On
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "On");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "On", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xLeft">Int32 xLeft</param>
        /// <param name="yTop">Int32 yTop</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Move(Int32 xLeft, Int32 yTop)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Move", xLeft, yTop);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Help()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Help");
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="on">bool on</param>
        /// <param name="callback">string callback</param>
        /// <param name="privateX">Int32 privateX</param>
        /// <param name="animation">optional object animation</param>
        /// <param name="customTeaser">optional object customTeaser</param>
        /// <param name="top">optional object top</param>
        /// <param name="left">optional object left</param>
        /// <param name="bottom">optional object bottom</param>
        /// <param name="right">optional object right</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 StartWizard(bool on, string callback, Int32 privateX, object animation, object customTeaser, object top, object left, object bottom, object right)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "StartWizard", new object[] { on, callback, privateX, animation, customTeaser, top, left, bottom, right });
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="on">bool on</param>
        /// <param name="callback">string callback</param>
        /// <param name="privateX">Int32 privateX</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 StartWizard(bool on, string callback, Int32 privateX)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "StartWizard", on, callback, privateX);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="on">bool on</param>
        /// <param name="callback">string callback</param>
        /// <param name="privateX">Int32 privateX</param>
        /// <param name="animation">optional object animation</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 StartWizard(bool on, string callback, Int32 privateX, object animation)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "StartWizard", on, callback, privateX, animation);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="on">bool on</param>
        /// <param name="callback">string callback</param>
        /// <param name="privateX">Int32 privateX</param>
        /// <param name="animation">optional object animation</param>
        /// <param name="customTeaser">optional object customTeaser</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 StartWizard(bool on, string callback, Int32 privateX, object animation, object customTeaser)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "StartWizard", new object[] { on, callback, privateX, animation, customTeaser });
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="on">bool on</param>
        /// <param name="callback">string callback</param>
        /// <param name="privateX">Int32 privateX</param>
        /// <param name="animation">optional object animation</param>
        /// <param name="customTeaser">optional object customTeaser</param>
        /// <param name="top">optional object top</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 StartWizard(bool on, string callback, Int32 privateX, object animation, object customTeaser, object top)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "StartWizard", new object[] { on, callback, privateX, animation, customTeaser, top });
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="on">bool on</param>
        /// <param name="callback">string callback</param>
        /// <param name="privateX">Int32 privateX</param>
        /// <param name="animation">optional object animation</param>
        /// <param name="customTeaser">optional object customTeaser</param>
        /// <param name="top">optional object top</param>
        /// <param name="left">optional object left</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 StartWizard(bool on, string callback, Int32 privateX, object animation, object customTeaser, object top, object left)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "StartWizard", new object[] { on, callback, privateX, animation, customTeaser, top, left });
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="on">bool on</param>
        /// <param name="callback">string callback</param>
        /// <param name="privateX">Int32 privateX</param>
        /// <param name="animation">optional object animation</param>
        /// <param name="customTeaser">optional object customTeaser</param>
        /// <param name="top">optional object top</param>
        /// <param name="left">optional object left</param>
        /// <param name="bottom">optional object bottom</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 StartWizard(bool on, string callback, Int32 privateX, object animation, object customTeaser, object top, object left, object bottom)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "StartWizard", new object[] { on, callback, privateX, animation, customTeaser, top, left, bottom });
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wizardID">Int32 wizardID</param>
        /// <param name="varfSuccess">bool varfSuccess</param>
        /// <param name="animation">optional object animation</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void EndWizard(Int32 wizardID, bool varfSuccess, object animation)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "EndWizard", wizardID, varfSuccess, animation);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wizardID">Int32 wizardID</param>
        /// <param name="varfSuccess">bool varfSuccess</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void EndWizard(Int32 wizardID, bool varfSuccess)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "EndWizard", wizardID, varfSuccess);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wizardID">Int32 wizardID</param>
        /// <param name="act">NetOffice.OfficeApi.Enums.MsoWizardActType act</param>
        /// <param name="animation">optional object animation</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ActivateWizard(Int32 wizardID, NetOffice.OfficeApi.Enums.MsoWizardActType act, object animation)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ActivateWizard", wizardID, act, animation);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wizardID">Int32 wizardID</param>
        /// <param name="act">NetOffice.OfficeApi.Enums.MsoWizardActType act</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ActivateWizard(Int32 wizardID, NetOffice.OfficeApi.Enums.MsoWizardActType act)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ActivateWizard", wizardID, act);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ResetTips()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ResetTips");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrAlertTitle">string bstrAlertTitle</param>
        /// <param name="bstrAlertText">string bstrAlertText</param>
        /// <param name="alb">NetOffice.OfficeApi.Enums.MsoAlertButtonType alb</param>
        /// <param name="alc">NetOffice.OfficeApi.Enums.MsoAlertIconType alc</param>
        /// <param name="ald">NetOffice.OfficeApi.Enums.MsoAlertDefaultType ald</param>
        /// <param name="alq">NetOffice.OfficeApi.Enums.MsoAlertCancelType alq</param>
        /// <param name="varfSysAlert">bool varfSysAlert</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 DoAlert(string bstrAlertTitle, string bstrAlertText, NetOffice.OfficeApi.Enums.MsoAlertButtonType alb, NetOffice.OfficeApi.Enums.MsoAlertIconType alc, NetOffice.OfficeApi.Enums.MsoAlertDefaultType ald, NetOffice.OfficeApi.Enums.MsoAlertCancelType alq, bool varfSysAlert)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DoAlert", new object[] { bstrAlertTitle, bstrAlertText, alb, alc, ald, alq, varfSysAlert });
        }

        #endregion

        #pragma warning restore
    }
}
