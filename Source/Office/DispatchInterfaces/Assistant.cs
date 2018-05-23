using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface Assistant 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface Assistant : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Top { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Left { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Balloon NewBalloon { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoBalloonErrorType BalloonError { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool Visible { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoAnimationType Animation { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool Reduced { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool AssistWithHelp { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool AssistWithWizards { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool AssistWithAlerts { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool MoveWhenInTheWay { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool Sounds { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool FeatureTips { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool MouseTips { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool KeyboardShortcutTips { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool HighPriorityTips { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool TipOfDay { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool GuessHelp { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool SearchWhenProgramming { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        string Item { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        string FileName { get; set; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        string Name { get; }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        bool On { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="xLeft">Int32 xLeft</param>
        /// <param name="yTop">Int32 yTop</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void Move(Int32 xLeft, Int32 yTop);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void Help();

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
        Int32 StartWizard(bool on, string callback, Int32 privateX, object animation, object customTeaser, object top, object left, object bottom, object right);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="on">bool on</param>
        /// <param name="callback">string callback</param>
        /// <param name="privateX">Int32 privateX</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        Int32 StartWizard(bool on, string callback, Int32 privateX);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="on">bool on</param>
        /// <param name="callback">string callback</param>
        /// <param name="privateX">Int32 privateX</param>
        /// <param name="animation">optional object animation</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        Int32 StartWizard(bool on, string callback, Int32 privateX, object animation);

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
        Int32 StartWizard(bool on, string callback, Int32 privateX, object animation, object customTeaser);

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
        Int32 StartWizard(bool on, string callback, Int32 privateX, object animation, object customTeaser, object top);

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
        Int32 StartWizard(bool on, string callback, Int32 privateX, object animation, object customTeaser, object top, object left);

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
        Int32 StartWizard(bool on, string callback, Int32 privateX, object animation, object customTeaser, object top, object left, object bottom);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wizardID">Int32 wizardID</param>
        /// <param name="varfSuccess">bool varfSuccess</param>
        /// <param name="animation">optional object animation</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void EndWizard(Int32 wizardID, bool varfSuccess, object animation);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wizardID">Int32 wizardID</param>
        /// <param name="varfSuccess">bool varfSuccess</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void EndWizard(Int32 wizardID, bool varfSuccess);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wizardID">Int32 wizardID</param>
        /// <param name="act">NetOffice.OfficeApi.Enums.MsoWizardActType act</param>
        /// <param name="animation">optional object animation</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void ActivateWizard(Int32 wizardID, NetOffice.OfficeApi.Enums.MsoWizardActType act, object animation);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="wizardID">Int32 wizardID</param>
        /// <param name="act">NetOffice.OfficeApi.Enums.MsoWizardActType act</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void ActivateWizard(Int32 wizardID, NetOffice.OfficeApi.Enums.MsoWizardActType act);

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        void ResetTips();

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
        Int32 DoAlert(string bstrAlertTitle, string bstrAlertText, NetOffice.OfficeApi.Enums.MsoAlertButtonType alb, NetOffice.OfficeApi.Enums.MsoAlertIconType alc, NetOffice.OfficeApi.Enums.MsoAlertDefaultType ald, NetOffice.OfficeApi.Enums.MsoAlertCancelType alq, bool varfSysAlert);

        #endregion
    }
}
