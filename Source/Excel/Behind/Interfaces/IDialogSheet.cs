using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface IDialogSheet 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IDialogSheet : COMObject, NetOffice.ExcelApi.IDialogSheet
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
                    _contractType = typeof(NetOffice.ExcelApi.IDialogSheet);
                return _contractType;
            }
        }
        private static Type _contractType;


        /// <summary>        /// Instance Type
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
                    _type = typeof(IDialogSheet);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IDialogSheet() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Application Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string CodeName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CodeName");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string _CodeName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "_CodeName");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "_CodeName", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Index
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Index");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string Name
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Name", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Next
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Next");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnDoubleClick
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnDoubleClick");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnDoubleClick", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnSheetActivate
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnSheetActivate");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnSheetActivate", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnSheetDeactivate
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnSheetDeactivate");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnSheetDeactivate", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.PageSetup PageSetup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PageSetup>(this, "PageSetup", typeof(NetOffice.ExcelApi.PageSetup));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Previous
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Previous");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ProtectContents
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ProtectContents");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ProtectDrawingObjects
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ProtectDrawingObjects");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ProtectionMode
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ProtectionMode");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ProtectScenarios
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ProtectScenarios");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlSheetVisibility Visible
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlSheetVisibility>(this, "Visible");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Visible", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Shapes Shapes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Shapes>(this, "Shapes", typeof(NetOffice.ExcelApi.Shapes));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool EnableCalculation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnableCalculation");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnableCalculation", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool DisplayAutomaticPageBreaks
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayAutomaticPageBreaks");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayAutomaticPageBreaks", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool EnableAutoFilter
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnableAutoFilter");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnableAutoFilter", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlEnableSelection EnableSelection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlEnableSelection>(this, "EnableSelection");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "EnableSelection", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool EnableOutlining
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnableOutlining");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnableOutlining", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool EnablePivotTable
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnablePivotTable");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnablePivotTable", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Names Names
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Names>(this, "Names", typeof(NetOffice.ExcelApi.Names));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string ScrollArea
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ScrollArea");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScrollArea", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.HPageBreaks HPageBreaks
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.HPageBreaks>(this, "HPageBreaks", typeof(NetOffice.ExcelApi.HPageBreaks));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.VPageBreaks VPageBreaks
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.VPageBreaks>(this, "VPageBreaks", typeof(NetOffice.ExcelApi.VPageBreaks));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.QueryTables QueryTables
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.QueryTables>(this, "QueryTables", typeof(NetOffice.ExcelApi.QueryTables));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool DisplayPageBreaks
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayPageBreaks");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayPageBreaks", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Comments Comments
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Comments>(this, "Comments", typeof(NetOffice.ExcelApi.Comments));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Hyperlinks Hyperlinks
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Hyperlinks>(this, "Hyperlinks", typeof(NetOffice.ExcelApi.Hyperlinks));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 _DisplayRightToLeft
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "_DisplayRightToLeft");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "_DisplayRightToLeft", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.AutoFilter AutoFilter
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.AutoFilter>(this, "AutoFilter", typeof(NetOffice.ExcelApi.AutoFilter));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool DisplayRightToLeft
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayRightToLeft");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayRightToLeft", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Scripts Scripts
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Scripts>(this, "Scripts", typeof(NetOffice.OfficeApi.Scripts));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object DefaultButton
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DefaultButton");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "DefaultButton", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.ExcelApi.DialogFrame DialogFrame
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.DialogFrame>(this, "DialogFrame", typeof(NetOffice.ExcelApi.DialogFrame));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Focus
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Focus");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Focus", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Tab Tab
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Tab>(this, "Tab", typeof(NetOffice.ExcelApi.Tab));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.MsoEnvelope MailEnvelope
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MsoEnvelope>(this, "MailEnvelope", typeof(NetOffice.OfficeApi.MsoEnvelope));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.CustomProperties CustomProperties
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.CustomProperties>(this, "CustomProperties", typeof(NetOffice.ExcelApi.CustomProperties));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.SmartTags SmartTags
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SmartTags>(this, "SmartTags", typeof(NetOffice.ExcelApi.SmartTags));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Protection Protection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Protection>(this, "Protection", typeof(NetOffice.ExcelApi.Protection));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool EnableFormatConditionsCalculation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnableFormatConditionsCalculation");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnableFormatConditionsCalculation", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Sort Sort
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sort>(this, "Sort", typeof(NetOffice.ExcelApi.Sort));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual Int32 PrintedCommentPages
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PrintedCommentPages");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Activate()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Activate");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="before">optional object before</param>
        /// <param name="after">optional object after</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Copy(object before, object after)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Copy", before, after);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Copy()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Copy");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="before">optional object before</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Copy(object before)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Copy", before);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Delete()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="before">optional object before</param>
        /// <param name="after">optional object after</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Move(object before, object after)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Move", before, after);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Move()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Move");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="before">optional object before</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Move(object before)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Move", before);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile, collate });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="prToFileName">optional object prToFileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile, collate, prToFileName });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _PrintOut()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_PrintOut");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _PrintOut(object from)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_PrintOut", from);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _PrintOut(object from, object to)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_PrintOut", from, to);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _PrintOut(object from, object to, object copies)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_PrintOut", from, to, copies);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _PrintOut(object from, object to, object copies, object preview)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_PrintOut", from, to, copies, preview);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _PrintOut(object from, object to, object copies, object preview, object activePrinter)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_PrintOut", new object[] { from, to, copies, preview, activePrinter });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="enableChanges">optional object enableChanges</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PrintPreview(object enableChanges)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PrintPreview", enableChanges);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PrintPreview()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PrintPreview");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        /// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Protect", new object[] { password, drawingObjects, contents, scenarios, userInterfaceOnly });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        /// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
        /// <param name="allowFormattingCells">optional object allowFormattingCells</param>
        /// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
        /// <param name="allowFormattingRows">optional object allowFormattingRows</param>
        /// <param name="allowInsertingColumns">optional object allowInsertingColumns</param>
        /// <param name="allowInsertingRows">optional object allowInsertingRows</param>
        /// <param name="allowInsertingHyperlinks">optional object allowInsertingHyperlinks</param>
        /// <param name="allowDeletingColumns">optional object allowDeletingColumns</param>
        /// <param name="allowDeletingRows">optional object allowDeletingRows</param>
        /// <param name="allowSorting">optional object allowSorting</param>
        /// <param name="allowFiltering">optional object allowFiltering</param>
        /// <param name="allowUsingPivotTables">optional object allowUsingPivotTables</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows, object allowSorting, object allowFiltering, object allowUsingPivotTables)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Protect", new object[] { password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows, allowSorting, allowFiltering, allowUsingPivotTables });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Protect()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Protect");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Protect(object password)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Protect", password);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Protect(object password, object drawingObjects)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Protect", password, drawingObjects);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Protect(object password, object drawingObjects, object contents)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Protect", password, drawingObjects, contents);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Protect(object password, object drawingObjects, object contents, object scenarios)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Protect", password, drawingObjects, contents, scenarios);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        /// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
        /// <param name="allowFormattingCells">optional object allowFormattingCells</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Protect", new object[] { password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        /// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
        /// <param name="allowFormattingCells">optional object allowFormattingCells</param>
        /// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Protect", new object[] { password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        /// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
        /// <param name="allowFormattingCells">optional object allowFormattingCells</param>
        /// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
        /// <param name="allowFormattingRows">optional object allowFormattingRows</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Protect", new object[] { password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        /// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
        /// <param name="allowFormattingCells">optional object allowFormattingCells</param>
        /// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
        /// <param name="allowFormattingRows">optional object allowFormattingRows</param>
        /// <param name="allowInsertingColumns">optional object allowInsertingColumns</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Protect", new object[] { password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        /// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
        /// <param name="allowFormattingCells">optional object allowFormattingCells</param>
        /// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
        /// <param name="allowFormattingRows">optional object allowFormattingRows</param>
        /// <param name="allowInsertingColumns">optional object allowInsertingColumns</param>
        /// <param name="allowInsertingRows">optional object allowInsertingRows</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Protect", new object[] { password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        /// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
        /// <param name="allowFormattingCells">optional object allowFormattingCells</param>
        /// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
        /// <param name="allowFormattingRows">optional object allowFormattingRows</param>
        /// <param name="allowInsertingColumns">optional object allowInsertingColumns</param>
        /// <param name="allowInsertingRows">optional object allowInsertingRows</param>
        /// <param name="allowInsertingHyperlinks">optional object allowInsertingHyperlinks</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Protect", new object[] { password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        /// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
        /// <param name="allowFormattingCells">optional object allowFormattingCells</param>
        /// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
        /// <param name="allowFormattingRows">optional object allowFormattingRows</param>
        /// <param name="allowInsertingColumns">optional object allowInsertingColumns</param>
        /// <param name="allowInsertingRows">optional object allowInsertingRows</param>
        /// <param name="allowInsertingHyperlinks">optional object allowInsertingHyperlinks</param>
        /// <param name="allowDeletingColumns">optional object allowDeletingColumns</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Protect", new object[] { password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        /// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
        /// <param name="allowFormattingCells">optional object allowFormattingCells</param>
        /// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
        /// <param name="allowFormattingRows">optional object allowFormattingRows</param>
        /// <param name="allowInsertingColumns">optional object allowInsertingColumns</param>
        /// <param name="allowInsertingRows">optional object allowInsertingRows</param>
        /// <param name="allowInsertingHyperlinks">optional object allowInsertingHyperlinks</param>
        /// <param name="allowDeletingColumns">optional object allowDeletingColumns</param>
        /// <param name="allowDeletingRows">optional object allowDeletingRows</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Protect", new object[] { password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        /// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
        /// <param name="allowFormattingCells">optional object allowFormattingCells</param>
        /// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
        /// <param name="allowFormattingRows">optional object allowFormattingRows</param>
        /// <param name="allowInsertingColumns">optional object allowInsertingColumns</param>
        /// <param name="allowInsertingRows">optional object allowInsertingRows</param>
        /// <param name="allowInsertingHyperlinks">optional object allowInsertingHyperlinks</param>
        /// <param name="allowDeletingColumns">optional object allowDeletingColumns</param>
        /// <param name="allowDeletingRows">optional object allowDeletingRows</param>
        /// <param name="allowSorting">optional object allowSorting</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows, object allowSorting)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Protect", new object[] { password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows, allowSorting });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        /// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
        /// <param name="allowFormattingCells">optional object allowFormattingCells</param>
        /// <param name="allowFormattingColumns">optional object allowFormattingColumns</param>
        /// <param name="allowFormattingRows">optional object allowFormattingRows</param>
        /// <param name="allowInsertingColumns">optional object allowInsertingColumns</param>
        /// <param name="allowInsertingRows">optional object allowInsertingRows</param>
        /// <param name="allowInsertingHyperlinks">optional object allowInsertingHyperlinks</param>
        /// <param name="allowDeletingColumns">optional object allowDeletingColumns</param>
        /// <param name="allowDeletingRows">optional object allowDeletingRows</param>
        /// <param name="allowSorting">optional object allowSorting</param>
        /// <param name="allowFiltering">optional object allowFiltering</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly, object allowFormattingCells, object allowFormattingColumns, object allowFormattingRows, object allowInsertingColumns, object allowInsertingRows, object allowInsertingHyperlinks, object allowDeletingColumns, object allowDeletingRows, object allowSorting, object allowFiltering)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Protect", new object[] { password, drawingObjects, contents, scenarios, userInterfaceOnly, allowFormattingCells, allowFormattingColumns, allowFormattingRows, allowInsertingColumns, allowInsertingRows, allowInsertingHyperlinks, allowDeletingColumns, allowDeletingRows, allowSorting, allowFiltering });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="addToMru">optional object addToMru</param>
        /// <param name="textCodepage">optional object textCodepage</param>
        /// <param name="textVisualLayout">optional object textVisualLayout</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage, object textVisualLayout)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage, textVisualLayout });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="addToMru">optional object addToMru</param>
        /// <param name="textCodepage">optional object textCodepage</param>
        /// <param name="textVisualLayout">optional object textVisualLayout</param>
        /// <param name="local">optional object local</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage, object textVisualLayout, object local)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage, textVisualLayout, local });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 SaveAs(string filename)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SaveAs", filename);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 SaveAs(string filename, object fileFormat)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SaveAs", filename, fileFormat);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 SaveAs(string filename, object fileFormat, object password)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SaveAs", filename, fileFormat, password);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 SaveAs(string filename, object fileFormat, object password, object writeResPassword)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SaveAs", filename, fileFormat, password, writeResPassword);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="addToMru">optional object addToMru</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="addToMru">optional object addToMru</param>
        /// <param name="textCodepage">optional object textCodepage</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="replace">optional object replace</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Select(object replace)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Select", replace);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Select()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Select");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Unprotect(object password)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Unprotect", password);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Unprotect()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Unprotect");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy29()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy29");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Arcs(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Arcs", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Arcs()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Arcs");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy31()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy31");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy32()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy32");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Buttons(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Buttons", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Buttons()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Buttons");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy34()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy34");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy36()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy36");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ChartObjects(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ChartObjects", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ChartObjects()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ChartObjects");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CheckBoxes(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CheckBoxes", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CheckBoxes()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "CheckBoxes");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="spellLang">optional object spellLang</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object spellLang)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest, spellLang);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 CheckSpelling()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CheckSpelling");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="customDictionary">optional object customDictionary</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 CheckSpelling(object customDictionary)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CheckSpelling", customDictionary);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 CheckSpelling(object customDictionary, object ignoreUppercase)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CheckSpelling", customDictionary, ignoreUppercase);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy40()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy40");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy41()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy41");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy42()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy42");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy43()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy43");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy44()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy44");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy45()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy45");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Drawings(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Drawings", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Drawings()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Drawings");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object DrawingObjects(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DrawingObjects", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object DrawingObjects()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DrawingObjects");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object DropDowns(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DropDowns", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object DropDowns()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DropDowns");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">object name</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Evaluate(object name)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Evaluate", name);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">object name</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object _Evaluate(object name)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_Evaluate", name);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy56()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy56");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 ResetAllPageBreaks()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ResetAllPageBreaks");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object GroupBoxes(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GroupBoxes", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object GroupBoxes()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GroupBoxes");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object GroupObjects(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GroupObjects", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object GroupObjects()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GroupObjects");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Labels(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Labels", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Labels()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Labels");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Lines(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Lines", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Lines()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Lines");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ListBoxes(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ListBoxes", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ListBoxes()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ListBoxes");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object OLEObjects(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "OLEObjects", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object OLEObjects()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "OLEObjects");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy65()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy65");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy66()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy66");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy67()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy67");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object OptionButtons(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "OptionButtons", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object OptionButtons()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "OptionButtons");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy69()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy69");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Ovals(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Ovals", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Ovals()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Ovals");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="destination">optional object destination</param>
        /// <param name="link">optional object link</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Paste(object destination, object link)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Paste", destination, link);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Paste()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Paste");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="destination">optional object destination</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Paste(object destination)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Paste", destination);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional object format</param>
        /// <param name="link">optional object link</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        /// <param name="iconFileName">optional object iconFileName</param>
        /// <param name="iconIndex">optional object iconIndex</param>
        /// <param name="iconLabel">optional object iconLabel</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PasteSpecial(object format, object link, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PasteSpecial", new object[] { format, link, displayAsIcon, iconFileName, iconIndex, iconLabel });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional object format</param>
        /// <param name="link">optional object link</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        /// <param name="iconFileName">optional object iconFileName</param>
        /// <param name="iconIndex">optional object iconIndex</param>
        /// <param name="iconLabel">optional object iconLabel</param>
        /// <param name="noHTMLFormatting">optional object noHTMLFormatting</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PasteSpecial(object format, object link, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object noHTMLFormatting)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PasteSpecial", new object[] { format, link, displayAsIcon, iconFileName, iconIndex, iconLabel, noHTMLFormatting });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PasteSpecial()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PasteSpecial");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional object format</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PasteSpecial(object format)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PasteSpecial", format);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional object format</param>
        /// <param name="link">optional object link</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PasteSpecial(object format, object link)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PasteSpecial", format, link);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional object format</param>
        /// <param name="link">optional object link</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PasteSpecial(object format, object link, object displayAsIcon)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PasteSpecial", format, link, displayAsIcon);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional object format</param>
        /// <param name="link">optional object link</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        /// <param name="iconFileName">optional object iconFileName</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PasteSpecial(object format, object link, object displayAsIcon, object iconFileName)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PasteSpecial", format, link, displayAsIcon, iconFileName);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional object format</param>
        /// <param name="link">optional object link</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        /// <param name="iconFileName">optional object iconFileName</param>
        /// <param name="iconIndex">optional object iconIndex</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PasteSpecial(object format, object link, object displayAsIcon, object iconFileName, object iconIndex)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PasteSpecial", new object[] { format, link, displayAsIcon, iconFileName, iconIndex });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Pictures(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Pictures", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Pictures()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Pictures");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy74()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy74");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy75()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy75");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy76()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy76");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Rectangles(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Rectangles", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Rectangles()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Rectangles");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy78()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy78");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy79()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy79");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ScrollBars(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ScrollBars", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object ScrollBars()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ScrollBars");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy82()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy82");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy83()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy83");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Spinners(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Spinners", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Spinners()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Spinners");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy85()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy85");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy86()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy86");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object TextBoxes(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextBoxes", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object TextBoxes()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "TextBoxes");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy88()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy88");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy89()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy89");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _Dummy90()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy90");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 ClearCircles()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ClearCircles");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 CircleInvalid()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CircleInvalid");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="prToFileName">optional object prToFileName</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile, collate, prToFileName });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PrintOut()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PrintOut");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PrintOut(object from)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PrintOut", from);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PrintOut(object from, object to)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PrintOut", from, to);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PrintOut(object from, object to, object copies)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PrintOut", from, to, copies);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PrintOut(object from, object to, object copies, object preview)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PrintOut", from, to, copies, preview);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PrintOut(object from, object to, object copies, object preview, object activePrinter)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PrintOut", new object[] { from, to, copies, preview, activePrinter });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile, collate });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="spellLang">optional object spellLang</param>
        /// <param name="ignoreFinalYaa">optional object ignoreFinalYaa</param>
        /// <param name="spellScript">optional object spellScript</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object spellLang, object ignoreFinalYaa, object spellScript)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_CheckSpelling", new object[] { customDictionary, ignoreUppercase, alwaysSuggest, spellLang, ignoreFinalYaa, spellScript });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _CheckSpelling()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_CheckSpelling");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="customDictionary">optional object customDictionary</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _CheckSpelling(object customDictionary)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_CheckSpelling", customDictionary);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _CheckSpelling(object customDictionary, object ignoreUppercase)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_CheckSpelling", customDictionary, ignoreUppercase);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="spellLang">optional object spellLang</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object spellLang)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest, spellLang);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="spellLang">optional object spellLang</param>
        /// <param name="ignoreFinalYaa">optional object ignoreFinalYaa</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object spellLang, object ignoreFinalYaa)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_CheckSpelling", new object[] { customDictionary, ignoreUppercase, alwaysSuggest, spellLang, ignoreFinalYaa });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object EditBoxes(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "EditBoxes", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object EditBoxes()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "EditBoxes");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="cancel">optional object cancel</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Hide(object cancel)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Hide", cancel);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Hide()
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Hide");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Show()
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Show");
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        /// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_Protect", new object[] { password, drawingObjects, contents, scenarios, userInterfaceOnly });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _Protect()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_Protect");
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _Protect(object password)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_Protect", password);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _Protect(object password, object drawingObjects)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_Protect", password, drawingObjects);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _Protect(object password, object drawingObjects, object contents)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_Protect", password, drawingObjects, contents);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _Protect(object password, object drawingObjects, object contents, object scenarios)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_Protect", password, drawingObjects, contents, scenarios);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="addToMru">optional object addToMru</param>
        /// <param name="textCodepage">optional object textCodepage</param>
        /// <param name="textVisualLayout">optional object textVisualLayout</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage, object textVisualLayout)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage, textVisualLayout });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _SaveAs(string filename)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_SaveAs", filename);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _SaveAs(string filename, object fileFormat)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_SaveAs", filename, fileFormat);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _SaveAs(string filename, object fileFormat, object password)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_SaveAs", filename, fileFormat, password);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _SaveAs(string filename, object fileFormat, object password, object writeResPassword)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_SaveAs", filename, fileFormat, password, writeResPassword);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="addToMru">optional object addToMru</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="addToMru">optional object addToMru</param>
        /// <param name="textCodepage">optional object textCodepage</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _SaveAs(string filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object addToMru, object textCodepage)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, addToMru, textCodepage });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional object format</param>
        /// <param name="link">optional object link</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        /// <param name="iconFileName">optional object iconFileName</param>
        /// <param name="iconIndex">optional object iconIndex</param>
        /// <param name="iconLabel">optional object iconLabel</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _PasteSpecial(object format, object link, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_PasteSpecial", new object[] { format, link, displayAsIcon, iconFileName, iconIndex, iconLabel });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _PasteSpecial()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_PasteSpecial");
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional object format</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _PasteSpecial(object format)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_PasteSpecial", format);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional object format</param>
        /// <param name="link">optional object link</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _PasteSpecial(object format, object link)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_PasteSpecial", format, link);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional object format</param>
        /// <param name="link">optional object link</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _PasteSpecial(object format, object link, object displayAsIcon)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_PasteSpecial", format, link, displayAsIcon);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional object format</param>
        /// <param name="link">optional object link</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        /// <param name="iconFileName">optional object iconFileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _PasteSpecial(object format, object link, object displayAsIcon, object iconFileName)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_PasteSpecial", format, link, displayAsIcon, iconFileName);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="format">optional object format</param>
        /// <param name="link">optional object link</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        /// <param name="iconFileName">optional object iconFileName</param>
        /// <param name="iconIndex">optional object iconIndex</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 _PasteSpecial(object format, object link, object displayAsIcon, object iconFileName, object iconIndex)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "_PasteSpecial", new object[] { format, link, displayAsIcon, iconFileName, iconIndex });
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual void _Dummy113()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy113");
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual void _Dummy114()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy114");
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual void _Dummy115()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Dummy115");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "__PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile, collate });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 __PrintOut()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "__PrintOut");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 __PrintOut(object from)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "__PrintOut", from);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 __PrintOut(object from, object to)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "__PrintOut", from, to);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 __PrintOut(object from, object to, object copies)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "__PrintOut", from, to, copies);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 __PrintOut(object from, object to, object copies, object preview)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "__PrintOut", from, to, copies, preview);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 __PrintOut(object from, object to, object copies, object preview, object activePrinter)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "__PrintOut", new object[] { from, to, copies, preview, activePrinter });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "__PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="openAfterPublish">optional object openAfterPublish</param>
        /// <param name="fixedFormatExtClassPtr">optional object fixedFormatExtClassPtr</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish, object fixedFormatExtClassPtr)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", new object[] { type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish, fixedFormatExtClassPtr });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", type);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", type, filename);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", type, filename, quality);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", type, filename, quality, includeDocProperties);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", new object[] { type, filename, quality, includeDocProperties, ignorePrintAreas });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        /// <param name="from">optional object from</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", new object[] { type, filename, quality, includeDocProperties, ignorePrintAreas, from });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", new object[] { type, filename, quality, includeDocProperties, ignorePrintAreas, from, to });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="openAfterPublish">optional object openAfterPublish</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Int32 ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ExportAsFixedFormat", new object[] { type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish });
        }

        #endregion

        #pragma warning restore
    }
}

