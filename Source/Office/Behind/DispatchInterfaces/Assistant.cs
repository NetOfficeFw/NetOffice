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
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
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
                return Factory.ExecuteInt32PropertyGet(this, "Top");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Top", value);
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
                return Factory.ExecuteInt32PropertyGet(this, "Left");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Left", value);
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
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Balloon>(this, "NewBalloon", typeof(NetOffice.OfficeApi.Balloon));
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
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoBalloonErrorType>(this, "BalloonError");
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
                return Factory.ExecuteBoolPropertyGet(this, "Visible");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Visible", value);
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
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoAnimationType>(this, "Animation");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "Animation", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "Reduced");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Reduced", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "AssistWithHelp");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "AssistWithHelp", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "AssistWithWizards");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "AssistWithWizards", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "AssistWithAlerts");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "AssistWithAlerts", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "MoveWhenInTheWay");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "MoveWhenInTheWay", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "Sounds");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Sounds", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "FeatureTips");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "FeatureTips", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "MouseTips");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "MouseTips", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "KeyboardShortcutTips");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "KeyboardShortcutTips", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "HighPriorityTips");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HighPriorityTips", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "TipOfDay");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "TipOfDay", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "GuessHelp");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "GuessHelp", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "SearchWhenProgramming");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "SearchWhenProgramming", value);
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
                return Factory.ExecuteStringPropertyGet(this, "Item");
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
                return Factory.ExecuteStringPropertyGet(this, "FileName");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "FileName", value);
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
                return Factory.ExecuteStringPropertyGet(this, "Name");
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
                return Factory.ExecuteBoolPropertyGet(this, "On");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "On", value);
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
            Factory.ExecuteMethod(this, "Move", xLeft, yTop);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Help()
        {
            Factory.ExecuteMethod(this, "Help");
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
            return Factory.ExecuteInt32MethodGet(this, "StartWizard", new object[] { on, callback, privateX, animation, customTeaser, top, left, bottom, right });
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
            return Factory.ExecuteInt32MethodGet(this, "StartWizard", on, callback, privateX);
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
            return Factory.ExecuteInt32MethodGet(this, "StartWizard", on, callback, privateX, animation);
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
            return Factory.ExecuteInt32MethodGet(this, "StartWizard", new object[] { on, callback, privateX, animation, customTeaser });
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
            return Factory.ExecuteInt32MethodGet(this, "StartWizard", new object[] { on, callback, privateX, animation, customTeaser, top });
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
            return Factory.ExecuteInt32MethodGet(this, "StartWizard", new object[] { on, callback, privateX, animation, customTeaser, top, left });
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
            return Factory.ExecuteInt32MethodGet(this, "StartWizard", new object[] { on, callback, privateX, animation, customTeaser, top, left, bottom });
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
            Factory.ExecuteMethod(this, "EndWizard", wizardID, varfSuccess, animation);
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
            Factory.ExecuteMethod(this, "EndWizard", wizardID, varfSuccess);
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
            Factory.ExecuteMethod(this, "ActivateWizard", wizardID, act, animation);
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
            Factory.ExecuteMethod(this, "ActivateWizard", wizardID, act);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ResetTips()
        {
            Factory.ExecuteMethod(this, "ResetTips");
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
            return Factory.ExecuteInt32MethodGet(this, "DoAlert", new object[] { bstrAlertTitle, bstrAlertText, alb, alc, ald, alq, varfSysAlert });
        }

        #endregion

        #pragma warning restore
    }
}
