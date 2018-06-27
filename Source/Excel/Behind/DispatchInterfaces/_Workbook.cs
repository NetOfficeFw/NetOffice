using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// _Workbook
    /// </summary>
    [SyntaxBypass]
    public class _Workbook_ : COMObject, NetOffice.ExcelApi._Workbook_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public _Workbook_() : base()
        {
            
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821660.aspx
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_Colors(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Colors", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_Colors(object index, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "Colors", index, value);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Colors
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821660.aspx </remarks>        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Colors")]
        public virtual object Colors(object index)
        {
            return get_Colors(index);
        }

        #endregion

        #region Methods

        #endregion
    }

    /// <summary>
    /// DispatchInterface _Workbook
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
    public class _Workbook : NetOffice.ExcelApi.Behind._Workbook_, NetOffice.ExcelApi._Workbook
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
                    _contractType = typeof(NetOffice.ExcelApi._Workbook);
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
                    _type = typeof(_Workbook);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public _Workbook() : base()
        {

        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835918.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840080.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198008.aspx </remarks>
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
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool AcceptLabelsInFormulas
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AcceptLabelsInFormulas");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AcceptLabelsInFormulas", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834923.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Chart ActiveChart
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Chart>(this, "ActiveChart", typeof(NetOffice.ExcelApi.Chart));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841181.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object ActiveSheet
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ActiveSheet");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string Author
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Author");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Author", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840067.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 AutoUpdateFrequency
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "AutoUpdateFrequency");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoUpdateFrequency", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193298.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool AutoUpdateSaveChanges
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoUpdateSaveChanges");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoUpdateSaveChanges", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821530.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 ChangeHistoryDuration
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ChangeHistoryDuration");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ChangeHistoryDuration", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197172.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object BuiltinDocumentProperties
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "BuiltinDocumentProperties");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821062.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Sheets Charts
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(this, "Charts", typeof(NetOffice.ExcelApi.Sheets));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195162.aspx </remarks>
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
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821660.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Colors
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Colors");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Colors", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835614.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBars CommandBars
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(this, "CommandBars", typeof(NetOffice.OfficeApi.CommandBars));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string Comments
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Comments");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Comments", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198339.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlSaveConflictResolution ConflictResolution
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlSaveConflictResolution>(this, "ConflictResolution");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ConflictResolution", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834401.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Container
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Container");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196337.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CreateBackup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CreateBackup");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834990.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object CustomDocumentProperties
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "CustomDocumentProperties");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193264.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Date1904
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Date1904");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Date1904", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.ExcelApi.Sheets DialogSheets
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(this, "DialogSheets", typeof(NetOffice.ExcelApi.Sheets));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834329.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.xlDisplayDrawingObjects DisplayDrawingObjects
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.xlDisplayDrawingObjects>(this, "DisplayDrawingObjects");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DisplayDrawingObjects", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840717.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlFileFormat FileFormat
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlFileFormat>(this, "FileFormat");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834975.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string FullName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FullName");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool HasMailer
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasMailer");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasMailer", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840238.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool HasPassword
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasPassword");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool HasRoutingSlip
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasRoutingSlip");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasRoutingSlip", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838249.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool IsAddin
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsAddin");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IsAddin", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string Keywords
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Keywords");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Keywords", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837965.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Mailer Mailer
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Mailer>(this, "Mailer", typeof(NetOffice.ExcelApi.Mailer));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.ExcelApi.Sheets Modules
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(this, "Modules", typeof(NetOffice.ExcelApi.Sheets));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839882.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool MultiUserEditing
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MultiUserEditing");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820899.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string Name
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195422.aspx </remarks>
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
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnSave
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnSave");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnSave", value);
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840974.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string Path
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Path");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836500.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool PersonalViewListSettings
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PersonalViewListSettings");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PersonalViewListSettings", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822649.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool PersonalViewPrintSettings
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PersonalViewPrintSettings");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PersonalViewPrintSettings", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198189.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool PrecisionAsDisplayed
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PrecisionAsDisplayed");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrecisionAsDisplayed", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838601.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ProtectStructure
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ProtectStructure");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193864.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ProtectWindows
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ProtectWindows");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840925.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ReadOnly
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ReadOnly");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196964.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ReadOnlyRecommended
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ReadOnlyRecommended");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ReadOnlyRecommended", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834665.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 RevisionNumber
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RevisionNumber");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Routed
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Routed");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.RoutingSlip RoutingSlip
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.RoutingSlip>(this, "RoutingSlip", typeof(NetOffice.ExcelApi.RoutingSlip));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196613.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Saved
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Saved");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Saved", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840667.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool SaveLinkValues
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SaveLinkValues");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SaveLinkValues", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197568.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Sheets Sheets
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(this, "Sheets", typeof(NetOffice.ExcelApi.Sheets));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839677.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ShowConflictHistory
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowConflictHistory");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowConflictHistory", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839039.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Styles Styles
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Styles>(this, "Styles", typeof(NetOffice.ExcelApi.Styles));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string Subject
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Subject");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Subject", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string Title
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Title");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Title", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193266.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool UpdateRemoteReferences
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UpdateRemoteReferences");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UpdateRemoteReferences", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool UserControl
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UserControl");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UserControl", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193788.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object UserStatus
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "UserStatus");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195531.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.CustomViews CustomViews
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.CustomViews>(this, "CustomViews", typeof(NetOffice.ExcelApi.CustomViews));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195152.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Windows Windows
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Windows>(this, "Windows", typeof(NetOffice.ExcelApi.Windows));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835542.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Sheets Worksheets
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(this, "Worksheets", typeof(NetOffice.ExcelApi.Sheets));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836228.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool WriteReserved
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "WriteReserved");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840737.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string WriteReservedBy
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "WriteReservedBy");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822819.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Sheets Excel4IntlMacroSheets
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(this, "Excel4IntlMacroSheets", typeof(NetOffice.ExcelApi.Sheets));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195645.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Sheets Excel4MacroSheets
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sheets>(this, "Excel4MacroSheets", typeof(NetOffice.ExcelApi.Sheets));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836472.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool TemplateRemoveExtData
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "TemplateRemoveExtData");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TemplateRemoveExtData", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194254.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool HighlightChangesOnScreen
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HighlightChangesOnScreen");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HighlightChangesOnScreen", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197016.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool KeepChangeHistory
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "KeepChangeHistory");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "KeepChangeHistory", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834301.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ListChangesOnNewSheet
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ListChangesOnNewSheet");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ListChangesOnNewSheet", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194737.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.VBIDEApi.VBProject VBProject
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBProject>(this, "VBProject");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840963.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool IsInplace
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsInplace");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838208.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.PublishObjects PublishObjects
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PublishObjects>(this, "PublishObjects", typeof(NetOffice.ExcelApi.PublishObjects));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834724.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.WebOptions WebOptions
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.WebOptions>(this, "WebOptions", typeof(NetOffice.ExcelApi.WebOptions));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.HTMLProject HTMLProject
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.HTMLProject>(this, "HTMLProject", typeof(NetOffice.OfficeApi.HTMLProject));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839554.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool EnvelopeVisible
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnvelopeVisible");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnvelopeVisible", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193512.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 CalculationVersion
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CalculationVersion");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822659.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool VBASigned
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "VBASigned");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool _ReadOnlyRecommended
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "_ReadOnlyRecommended");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196322.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool ShowPivotTableFieldList
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowPivotTableFieldList");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowPivotTableFieldList", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839021.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlUpdateLinks UpdateLinks
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlUpdateLinks>(this, "UpdateLinks");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "UpdateLinks", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193225.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool EnableAutoRecover
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EnableAutoRecover");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EnableAutoRecover", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841017.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool RemovePersonalInformation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "RemovePersonalInformation");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RemovePersonalInformation", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821089.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual string FullNameURLEncoded
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FullNameURLEncoded");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821529.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual string Password
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Password");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Password", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837767.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual string WritePassword
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "WritePassword");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WritePassword", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839579.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual string PasswordEncryptionProvider
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PasswordEncryptionProvider");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195464.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual string PasswordEncryptionAlgorithm
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PasswordEncryptionAlgorithm");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195381.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PasswordEncryptionKeyLength
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PasswordEncryptionKeyLength");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820819.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool PasswordEncryptionFileProperties
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PasswordEncryptionFileProperties");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.SmartTagOptions SmartTagOptions
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SmartTagOptions>(this, "SmartTagOptions", typeof(NetOffice.ExcelApi.SmartTagOptions));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840697.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Permission Permission
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Permission>(this, "Permission", typeof(NetOffice.OfficeApi.Permission));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835236.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.SharedWorkspace SharedWorkspace
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SharedWorkspace>(this, "SharedWorkspace", typeof(NetOffice.OfficeApi.SharedWorkspace));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192923.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Sync Sync
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Sync>(this, "Sync", typeof(NetOffice.OfficeApi.Sync));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838260.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.XmlNamespaces XmlNamespaces
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.XmlNamespaces>(this, "XmlNamespaces", typeof(NetOffice.ExcelApi.XmlNamespaces));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838975.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.XmlMaps XmlMaps
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.XmlMaps>(this, "XmlMaps", typeof(NetOffice.ExcelApi.XmlMaps));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194561.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.SmartDocument SmartDocument
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartDocument>(this, "SmartDocument", typeof(NetOffice.OfficeApi.SmartDocument));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838205.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.DocumentLibraryVersions DocumentLibraryVersions
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DocumentLibraryVersions>(this, "DocumentLibraryVersions", typeof(NetOffice.OfficeApi.DocumentLibraryVersions));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837429.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual bool InactiveListBorderVisible
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "InactiveListBorderVisible");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InactiveListBorderVisible", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838435.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual bool DisplayInkComments
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayInkComments");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayInkComments", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837152.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.MetaProperties ContentTypeProperties
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MetaProperties>(this, "ContentTypeProperties", typeof(NetOffice.OfficeApi.MetaProperties));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836773.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Connections Connections
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Connections>(this, "Connections", typeof(NetOffice.ExcelApi.Connections));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838073.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.SignatureSet Signatures
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SignatureSet>(this, "Signatures", typeof(NetOffice.OfficeApi.SignatureSet));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194489.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.ServerPolicy ServerPolicy
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ServerPolicy>(this, "ServerPolicy", typeof(NetOffice.OfficeApi.ServerPolicy));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195426.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.DocumentInspectors DocumentInspectors
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.DocumentInspectors>(this, "DocumentInspectors", typeof(NetOffice.OfficeApi.DocumentInspectors));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195818.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.ServerViewableItems ServerViewableItems
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ServerViewableItems>(this, "ServerViewableItems", typeof(NetOffice.ExcelApi.ServerViewableItems));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837756.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.TableStyles TableStyles
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.TableStyles>(this, "TableStyles", typeof(NetOffice.ExcelApi.TableStyles));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195934.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object DefaultTableStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DefaultTableStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "DefaultTableStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835624.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object DefaultPivotTableStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DefaultPivotTableStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "DefaultPivotTableStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836165.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool CheckCompatibility
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CheckCompatibility");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CheckCompatibility", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838063.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool HasVBProject
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasVBProject");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838448.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CustomXMLParts CustomXMLParts
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CustomXMLParts>(this, "CustomXMLParts", typeof(NetOffice.OfficeApi.CustomXMLParts));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820907.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool Final
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Final");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Final", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196847.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Research Research
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Research>(this, "Research", typeof(NetOffice.ExcelApi.Research));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194072.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.OfficeTheme Theme
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.OfficeTheme>(this, "Theme", typeof(NetOffice.OfficeApi.OfficeTheme));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834991.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool Excel8CompatibilityMode
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Excel8CompatibilityMode");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837960.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool ConnectionsDisabled
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ConnectionsDisabled");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835280.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool ShowPivotChartActiveFields
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowPivotChartActiveFields");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowPivotChartActiveFields", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839003.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.IconSets IconSets
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.IconSets>(this, "IconSets", typeof(NetOffice.ExcelApi.IconSets));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194147.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual string EncryptionProvider
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EncryptionProvider");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EncryptionProvider", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839440.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool DoNotPromptForConvert
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DoNotPromptForConvert");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DoNotPromptForConvert", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823189.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool ForceFullCalculation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ForceFullCalculation");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ForceFullCalculation", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194925.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual NetOffice.ExcelApi.SlicerCaches SlicerCaches
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SlicerCaches>(this, "SlicerCaches", typeof(NetOffice.ExcelApi.SlicerCaches));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839464.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Slicer ActiveSlicer
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Slicer>(this, "ActiveSlicer", typeof(NetOffice.ExcelApi.Slicer));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193862.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object DefaultSlicerStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DefaultSlicerStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "DefaultSlicerStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838425.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual Int32 AccuracyVersion
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "AccuracyVersion");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AccuracyVersion", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229542.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        public virtual bool CaseSensitive
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CaseSensitive");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231362.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        public virtual bool UseWholeCellCriteria
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UseWholeCellCriteria");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230772.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        public virtual bool UseWildcards
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UseWildcards");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231277.aspx </remarks>
        [SupportByVersion("Excel", 15, 16), ProxyResult]
        public virtual object PivotTables
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "PivotTables");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228926.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.Model Model
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Model>(this, "Model", typeof(NetOffice.ExcelApi.Model));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227452.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        public virtual bool ChartDataPointTrack
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ChartDataPointTrack");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ChartDataPointTrack", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230214.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object DefaultTimelineStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "DefaultTimelineStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "DefaultTimelineStyle", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821837.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Activate()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Activate");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193344.aspx </remarks>
        /// <param name="mode">NetOffice.ExcelApi.Enums.XlFileAccess mode</param>
        /// <param name="writePassword">optional object writePassword</param>
        /// <param name="notify">optional object notify</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ChangeFileAccess(NetOffice.ExcelApi.Enums.XlFileAccess mode, object writePassword, object notify)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChangeFileAccess", mode, writePassword, notify);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193344.aspx </remarks>
        /// <param name="mode">NetOffice.ExcelApi.Enums.XlFileAccess mode</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ChangeFileAccess(NetOffice.ExcelApi.Enums.XlFileAccess mode)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChangeFileAccess", mode);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193344.aspx </remarks>
        /// <param name="mode">NetOffice.ExcelApi.Enums.XlFileAccess mode</param>
        /// <param name="writePassword">optional object writePassword</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ChangeFileAccess(NetOffice.ExcelApi.Enums.XlFileAccess mode, object writePassword)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChangeFileAccess", mode, writePassword);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836537.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="newName">string newName</param>
        /// <param name="type">optional NetOffice.ExcelApi.Enums.XlLinkType Type = 1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ChangeLink(string name, string newName, object type)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChangeLink", name, newName, type);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836537.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="newName">string newName</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ChangeLink(string name, string newName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChangeLink", name, newName);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838613.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="routeWorkbook">optional object routeWorkbook</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Close(object saveChanges, object filename, object routeWorkbook)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Close", saveChanges, filename, routeWorkbook);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838613.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Close()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Close");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838613.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Close(object saveChanges)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Close", saveChanges);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838613.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="filename">optional object filename</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Close(object saveChanges, object filename)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Close", saveChanges, filename);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839565.aspx </remarks>
        /// <param name="numberFormat">string numberFormat</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void DeleteNumberFormat(string numberFormat)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteNumberFormat", numberFormat);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836762.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ExclusiveAccess()
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ExclusiveAccess");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836208.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ForwardMailer()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ForwardMailer");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192971.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="linkInfo">NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo</param>
        /// <param name="type">optional object type</param>
        /// <param name="editionRef">optional object editionRef</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object LinkInfo(string name, NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo, object type, object editionRef)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LinkInfo", name, linkInfo, type, editionRef);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192971.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="linkInfo">NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object LinkInfo(string name, NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LinkInfo", name, linkInfo);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192971.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="linkInfo">NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo</param>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object LinkInfo(string name, NetOffice.ExcelApi.Enums.XlLinkInfo linkInfo, object type)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LinkInfo", name, linkInfo, type);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821922.aspx </remarks>
        /// <param name="type">optional object type</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object LinkSources(object type)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LinkSources", type);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821922.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object LinkSources()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LinkSources");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196693.aspx </remarks>
        /// <param name="filename">object filename</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void MergeWorkbook(object filename)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "MergeWorkbook", filename);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838378.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Window NewWindow()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Window>(this, "NewWindow", typeof(NetOffice.ExcelApi.Window));
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839052.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="readOnly">optional object readOnly</param>
        /// <param name="type">optional object type</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void OpenLinks(string name, object readOnly, object type)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "OpenLinks", name, readOnly, type);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839052.aspx </remarks>
        /// <param name="name">string name</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void OpenLinks(string name)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "OpenLinks", name);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839052.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="readOnly">optional object readOnly</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void OpenLinks(string name, object readOnly)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "OpenLinks", name, readOnly);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193549.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.PivotCaches PivotCaches()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotCaches>(this, "PivotCaches", typeof(NetOffice.ExcelApi.PivotCaches));
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821844.aspx </remarks>
        /// <param name="destName">optional object destName</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Post(object destName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Post", destName);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821844.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Post()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Post");
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
        public virtual void _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile, collate });
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
        public virtual void _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile, collate, prToFileName });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _PrintOut()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_PrintOut");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _PrintOut(object from)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_PrintOut", from);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void _PrintOut(object from, object to)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_PrintOut", from, to);
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
        public virtual void _PrintOut(object from, object to, object copies)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_PrintOut", from, to, copies);
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
        public virtual void _PrintOut(object from, object to, object copies, object preview)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_PrintOut", from, to, copies, preview);
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
        public virtual void _PrintOut(object from, object to, object copies, object preview, object activePrinter)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_PrintOut", new object[] { from, to, copies, preview, activePrinter });
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
        public virtual void _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193068.aspx </remarks>
        /// <param name="enableChanges">optional object enableChanges</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintPreview(object enableChanges)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintPreview", enableChanges);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193068.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintPreview()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintPreview");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193800.aspx </remarks>
        /// <param name="password">optional object password</param>
        /// <param name="structure">optional object structure</param>
        /// <param name="windows">optional object windows</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Protect(object password, object structure, object windows)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Protect", password, structure, windows);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193800.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Protect()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Protect");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193800.aspx </remarks>
        /// <param name="password">optional object password</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Protect(object password)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Protect", password);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193800.aspx </remarks>
        /// <param name="password">optional object password</param>
        /// <param name="structure">optional object structure</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Protect(object password, object structure)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Protect", password, structure);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="sharingPassword">optional object sharingPassword</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object sharingPassword)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ProtectSharing", new object[] { filename, password, writeResPassword, readOnlyRecommended, createBackup, sharingPassword });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="sharingPassword">optional object sharingPassword</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object sharingPassword, object fileFormat)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ProtectSharing", new object[] { filename, password, writeResPassword, readOnlyRecommended, createBackup, sharingPassword, fileFormat });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ProtectSharing()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ProtectSharing");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ProtectSharing(object filename)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ProtectSharing", filename);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ProtectSharing(object filename, object password)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ProtectSharing", filename, password);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ProtectSharing(object filename, object password, object writeResPassword)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ProtectSharing", filename, password, writeResPassword);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ProtectSharing", filename, password, writeResPassword, readOnlyRecommended);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195383.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ProtectSharing", new object[] { filename, password, writeResPassword, readOnlyRecommended, createBackup });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838648.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void RefreshAll()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RefreshAll");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820902.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Reply()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Reply");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838788.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ReplyAll()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ReplyAll");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840747.aspx </remarks>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void RemoveUser(Int32 index)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveUser", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Route()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Route");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835203.aspx </remarks>
        /// <param name="which">NetOffice.ExcelApi.Enums.XlRunAutoMacro which</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void RunAutoMacros(NetOffice.ExcelApi.Enums.XlRunAutoMacro which)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RunAutoMacros", which);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197585.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Save()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Save");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        /// <param name="conflictResolution">optional object conflictResolution</param>
        /// <param name="addToMru">optional object addToMru</param>
        /// <param name="textCodepage">optional object textCodepage</param>
        /// <param name="textVisualLayout">optional object textVisualLayout</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage, object textVisualLayout)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, textVisualLayout });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        /// <param name="conflictResolution">optional object conflictResolution</param>
        /// <param name="addToMru">optional object addToMru</param>
        /// <param name="textCodepage">optional object textCodepage</param>
        /// <param name="textVisualLayout">optional object textVisualLayout</param>
        /// <param name="local">optional object local</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage, object textVisualLayout, object local)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, textVisualLayout, local });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SaveAs()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SaveAs(object filename)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", filename);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SaveAs(object filename, object fileFormat)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", filename, fileFormat);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SaveAs(object filename, object fileFormat, object password)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", filename, fileFormat, password);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SaveAs(object filename, object fileFormat, object password, object writeResPassword)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", filename, fileFormat, password, writeResPassword);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        /// <param name="conflictResolution">optional object conflictResolution</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        /// <param name="conflictResolution">optional object conflictResolution</param>
        /// <param name="addToMru">optional object addToMru</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841185.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        /// <param name="conflictResolution">optional object conflictResolution</param>
        /// <param name="addToMru">optional object addToMru</param>
        /// <param name="textCodepage">optional object textCodepage</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835014.aspx </remarks>
        /// <param name="filename">optional object filename</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SaveCopyAs(object filename)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs", filename);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835014.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SaveCopyAs()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveCopyAs");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821053.aspx </remarks>
        /// <param name="recipients">object recipients</param>
        /// <param name="subject">optional object subject</param>
        /// <param name="returnReceipt">optional object returnReceipt</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SendMail(object recipients, object subject, object returnReceipt)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SendMail", recipients, subject, returnReceipt);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821053.aspx </remarks>
        /// <param name="recipients">object recipients</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SendMail(object recipients)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SendMail", recipients);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821053.aspx </remarks>
        /// <param name="recipients">object recipients</param>
        /// <param name="subject">optional object subject</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SendMail(object recipients, object subject)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SendMail", recipients, subject);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840261.aspx </remarks>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="priority">optional NetOffice.ExcelApi.Enums.XlPriority Priority = -4143</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SendMailer(object fileFormat, object priority)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SendMailer", fileFormat, priority);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840261.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SendMailer()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SendMailer");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840261.aspx </remarks>
        /// <param name="fileFormat">optional object fileFormat</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SendMailer(object fileFormat)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SendMailer", fileFormat);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838177.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="procedure">optional object procedure</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SetLinkOnData(string name, object procedure)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetLinkOnData", name, procedure);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838177.aspx </remarks>
        /// <param name="name">string name</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SetLinkOnData(string name)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetLinkOnData", name);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196695.aspx </remarks>
        /// <param name="password">optional object password</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Unprotect(object password)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Unprotect", password);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196695.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Unprotect()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Unprotect");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840639.aspx </remarks>
        /// <param name="sharingPassword">optional object sharingPassword</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void UnprotectSharing(object sharingPassword)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UnprotectSharing", sharingPassword);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840639.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void UnprotectSharing()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UnprotectSharing");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840979.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void UpdateFromFile()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateFromFile");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195741.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="type">optional object type</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void UpdateLink(object name, object type)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateLink", name, type);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195741.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void UpdateLink()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateLink");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195741.aspx </remarks>
        /// <param name="name">optional object name</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void UpdateLink(object name)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateLink", name);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837763.aspx </remarks>
        /// <param name="when">optional object when</param>
        /// <param name="who">optional object who</param>
        /// <param name="where">optional object where</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void HighlightChangesOptions(object when, object who, object where)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "HighlightChangesOptions", when, who, where);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837763.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void HighlightChangesOptions()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "HighlightChangesOptions");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837763.aspx </remarks>
        /// <param name="when">optional object when</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void HighlightChangesOptions(object when)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "HighlightChangesOptions", when);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837763.aspx </remarks>
        /// <param name="when">optional object when</param>
        /// <param name="who">optional object who</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void HighlightChangesOptions(object when, object who)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "HighlightChangesOptions", when, who);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834662.aspx </remarks>
        /// <param name="days">Int32 days</param>
        /// <param name="sharingPassword">optional object sharingPassword</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PurgeChangeHistoryNow(Int32 days, object sharingPassword)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PurgeChangeHistoryNow", days, sharingPassword);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834662.aspx </remarks>
        /// <param name="days">Int32 days</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PurgeChangeHistoryNow(Int32 days)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PurgeChangeHistoryNow", days);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835613.aspx </remarks>
        /// <param name="when">optional object when</param>
        /// <param name="who">optional object who</param>
        /// <param name="where">optional object where</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void AcceptAllChanges(object when, object who, object where)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AcceptAllChanges", when, who, where);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835613.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void AcceptAllChanges()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AcceptAllChanges");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835613.aspx </remarks>
        /// <param name="when">optional object when</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void AcceptAllChanges(object when)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AcceptAllChanges", when);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835613.aspx </remarks>
        /// <param name="when">optional object when</param>
        /// <param name="who">optional object who</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void AcceptAllChanges(object when, object who)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AcceptAllChanges", when, who);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837407.aspx </remarks>
        /// <param name="when">optional object when</param>
        /// <param name="who">optional object who</param>
        /// <param name="where">optional object where</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void RejectAllChanges(object when, object who, object where)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RejectAllChanges", when, who, where);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837407.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void RejectAllChanges()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RejectAllChanges");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837407.aspx </remarks>
        /// <param name="when">optional object when</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void RejectAllChanges(object when)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RejectAllChanges", when);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837407.aspx </remarks>
        /// <param name="when">optional object when</param>
        /// <param name="who">optional object who</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void RejectAllChanges(object when, object who)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RejectAllChanges", when, who);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        /// <param name="backgroundQuery">optional object backgroundQuery</param>
        /// <param name="optimizeCache">optional object optimizeCache</param>
        /// <param name="pageFieldOrder">optional object pageFieldOrder</param>
        /// <param name="pageFieldWrapCount">optional object pageFieldWrapCount</param>
        /// <param name="readData">optional object readData</param>
        /// <param name="connection">optional object connection</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount, object readData, object connection)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount, readData, connection });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PivotTableWizard()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PivotTableWizard");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PivotTableWizard(object sourceType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PivotTableWizard", sourceType);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PivotTableWizard(object sourceType, object sourceData)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PivotTableWizard", sourceType, sourceData);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PivotTableWizard(object sourceType, object sourceData, object tableDestination)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PivotTableWizard", sourceType, sourceData, tableDestination);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PivotTableWizard", sourceType, sourceData, tableDestination, tableName);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        /// <param name="backgroundQuery">optional object backgroundQuery</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        /// <param name="backgroundQuery">optional object backgroundQuery</param>
        /// <param name="optimizeCache">optional object optimizeCache</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        /// <param name="backgroundQuery">optional object backgroundQuery</param>
        /// <param name="optimizeCache">optional object optimizeCache</param>
        /// <param name="pageFieldOrder">optional object pageFieldOrder</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        /// <param name="backgroundQuery">optional object backgroundQuery</param>
        /// <param name="optimizeCache">optional object optimizeCache</param>
        /// <param name="pageFieldOrder">optional object pageFieldOrder</param>
        /// <param name="pageFieldWrapCount">optional object pageFieldWrapCount</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sourceType">optional object sourceType</param>
        /// <param name="sourceData">optional object sourceData</param>
        /// <param name="tableDestination">optional object tableDestination</param>
        /// <param name="tableName">optional object tableName</param>
        /// <param name="rowGrand">optional object rowGrand</param>
        /// <param name="columnGrand">optional object columnGrand</param>
        /// <param name="saveData">optional object saveData</param>
        /// <param name="hasAutoFormat">optional object hasAutoFormat</param>
        /// <param name="autoPage">optional object autoPage</param>
        /// <param name="reserved">optional object reserved</param>
        /// <param name="backgroundQuery">optional object backgroundQuery</param>
        /// <param name="optimizeCache">optional object optimizeCache</param>
        /// <param name="pageFieldOrder">optional object pageFieldOrder</param>
        /// <param name="pageFieldWrapCount">optional object pageFieldWrapCount</param>
        /// <param name="readData">optional object readData</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand, object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery, object optimizeCache, object pageFieldOrder, object pageFieldWrapCount, object readData)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PivotTableWizard", new object[] { sourceType, sourceData, tableDestination, tableName, rowGrand, columnGrand, saveData, hasAutoFormat, autoPage, reserved, backgroundQuery, optimizeCache, pageFieldOrder, pageFieldWrapCount, readData });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194697.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ResetColors()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ResetColors");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
        /// <param name="address">string address</param>
        /// <param name="subAddress">optional object subAddress</param>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="addHistory">optional object addHistory</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        /// <param name="method">optional object method</param>
        /// <param name="headerInfo">optional object headerInfo</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method, object headerInfo)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", new object[] { address, subAddress, newWindow, addHistory, extraInfo, method, headerInfo });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
        /// <param name="address">string address</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void FollowHyperlink(string address)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
        /// <param name="address">string address</param>
        /// <param name="subAddress">optional object subAddress</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void FollowHyperlink(string address, object subAddress)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address, subAddress);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
        /// <param name="address">string address</param>
        /// <param name="subAddress">optional object subAddress</param>
        /// <param name="newWindow">optional object newWindow</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void FollowHyperlink(string address, object subAddress, object newWindow)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address, subAddress, newWindow);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
        /// <param name="address">string address</param>
        /// <param name="subAddress">optional object subAddress</param>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="addHistory">optional object addHistory</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address, subAddress, newWindow, addHistory);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
        /// <param name="address">string address</param>
        /// <param name="subAddress">optional object subAddress</param>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="addHistory">optional object addHistory</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", new object[] { address, subAddress, newWindow, addHistory, extraInfo });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839280.aspx </remarks>
        /// <param name="address">string address</param>
        /// <param name="subAddress">optional object subAddress</param>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="addHistory">optional object addHistory</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        /// <param name="method">optional object method</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", new object[] { address, subAddress, newWindow, addHistory, extraInfo, method });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194282.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void AddToFavorites()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AddToFavorites");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="prToFileName">optional object prToFileName</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile, collate, prToFileName });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="prToFileName">optional object prToFileName</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName, object ignorePrintAreas)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile, collate, prToFileName, ignorePrintAreas });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        /// <param name="from">optional object from</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object from)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", from);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object from, object to)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", from, to);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object from, object to, object copies)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", from, to, copies);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object from, object to, object copies, object preview)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", from, to, copies, preview);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object from, object to, object copies, object preview, object activePrinter)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { from, to, copies, preview, activePrinter });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196840.aspx </remarks>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="preview">optional object preview</param>
        /// <param name="activePrinter">optional object activePrinter</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile, collate });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195831.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void WebPagePreview()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "WebPagePreview");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839234.aspx </remarks>
        /// <param name="encoding">NetOffice.OfficeApi.Enums.MsoEncoding encoding</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ReloadAs(NetOffice.OfficeApi.Enums.MsoEncoding encoding)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ReloadAs", encoding);
        }

        /// <summary>
        /// SupportByVersion Excel 9
        /// </summary>
        /// <param name="unused">Int32 unused</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9)]
        public virtual void Dummy1(Int32 unused)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Dummy1", unused);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="s">string s</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void sblt(string s)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "sblt", s);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="structure">optional object structure</param>
        /// <param name="windows">optional object windows</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void _Protect(object password, object structure, object windows)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Protect", password, structure, windows);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void _Protect()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Protect");
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void _Protect(object password)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Protect", password);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="structure">optional object structure</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void _Protect(object password, object structure)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_Protect", password, structure);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        /// <param name="conflictResolution">optional object conflictResolution</param>
        /// <param name="addToMru">optional object addToMru</param>
        /// <param name="textCodepage">optional object textCodepage</param>
        /// <param name="textVisualLayout">optional object textVisualLayout</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage, object textVisualLayout)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, textVisualLayout });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void _SaveAs()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_SaveAs");
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void _SaveAs(object filename)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_SaveAs", filename);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void _SaveAs(object filename, object fileFormat)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_SaveAs", filename, fileFormat);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void _SaveAs(object filename, object fileFormat, object password)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_SaveAs", filename, fileFormat, password);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void _SaveAs(object filename, object fileFormat, object password, object writeResPassword)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_SaveAs", filename, fileFormat, password, writeResPassword);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        /// <param name="conflictResolution">optional object conflictResolution</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        /// <param name="conflictResolution">optional object conflictResolution</param>
        /// <param name="addToMru">optional object addToMru</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="fileFormat">optional object fileFormat</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="accessMode">optional NetOffice.ExcelApi.Enums.XlSaveAsAccessMode AccessMode = 1</param>
        /// <param name="conflictResolution">optional object conflictResolution</param>
        /// <param name="addToMru">optional object addToMru</param>
        /// <param name="textCodepage">optional object textCodepage</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object accessMode, object conflictResolution, object addToMru, object textCodepage)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_SaveAs", new object[] { filename, fileFormat, password, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="calcid">Int32 calcid</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void Dummy17(Int32 calcid)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Dummy17", calcid);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194915.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlLinkType type</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void BreakLink(string name, NetOffice.ExcelApi.Enums.XlLinkType type)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "BreakLink", name, type);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void Dummy16()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Dummy16");
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841145.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="comments">optional object comments</param>
        /// <param name="makePublic">optional object makePublic</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void CheckIn(object saveChanges, object comments, object makePublic)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn", saveChanges, comments, makePublic);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841145.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void CheckIn()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn");
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841145.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void CheckIn(object saveChanges)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn", saveChanges);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841145.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="comments">optional object comments</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void CheckIn(object saveChanges, object comments)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn", saveChanges, comments);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194456.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual bool CanCheckIn()
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CanCheckIn");
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx </remarks>
        /// <param name="recipients">optional object recipients</param>
        /// <param name="subject">optional object subject</param>
        /// <param name="showMessage">optional object showMessage</param>
        /// <param name="includeAttachment">optional object includeAttachment</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void SendForReview(object recipients, object subject, object showMessage, object includeAttachment)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SendForReview", recipients, subject, showMessage, includeAttachment);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void SendForReview()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SendForReview");
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx </remarks>
        /// <param name="recipients">optional object recipients</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void SendForReview(object recipients)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SendForReview", recipients);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx </remarks>
        /// <param name="recipients">optional object recipients</param>
        /// <param name="subject">optional object subject</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void SendForReview(object recipients, object subject)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SendForReview", recipients, subject);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196626.aspx </remarks>
        /// <param name="recipients">optional object recipients</param>
        /// <param name="subject">optional object subject</param>
        /// <param name="showMessage">optional object showMessage</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void SendForReview(object recipients, object subject, object showMessage)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SendForReview", recipients, subject, showMessage);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821626.aspx </remarks>
        /// <param name="showMessage">optional object showMessage</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void ReplyWithChanges(object showMessage)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ReplyWithChanges", showMessage);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821626.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void ReplyWithChanges()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ReplyWithChanges");
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839207.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void EndReview()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "EndReview");
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx </remarks>
        /// <param name="passwordEncryptionProvider">optional object passwordEncryptionProvider</param>
        /// <param name="passwordEncryptionAlgorithm">optional object passwordEncryptionAlgorithm</param>
        /// <param name="passwordEncryptionKeyLength">optional object passwordEncryptionKeyLength</param>
        /// <param name="passwordEncryptionFileProperties">optional object passwordEncryptionFileProperties</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void SetPasswordEncryptionOptions(object passwordEncryptionProvider, object passwordEncryptionAlgorithm, object passwordEncryptionKeyLength, object passwordEncryptionFileProperties)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetPasswordEncryptionOptions", passwordEncryptionProvider, passwordEncryptionAlgorithm, passwordEncryptionKeyLength, passwordEncryptionFileProperties);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void SetPasswordEncryptionOptions()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetPasswordEncryptionOptions");
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx </remarks>
        /// <param name="passwordEncryptionProvider">optional object passwordEncryptionProvider</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void SetPasswordEncryptionOptions(object passwordEncryptionProvider)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetPasswordEncryptionOptions", passwordEncryptionProvider);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx </remarks>
        /// <param name="passwordEncryptionProvider">optional object passwordEncryptionProvider</param>
        /// <param name="passwordEncryptionAlgorithm">optional object passwordEncryptionAlgorithm</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void SetPasswordEncryptionOptions(object passwordEncryptionProvider, object passwordEncryptionAlgorithm)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetPasswordEncryptionOptions", passwordEncryptionProvider, passwordEncryptionAlgorithm);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196907.aspx </remarks>
        /// <param name="passwordEncryptionProvider">optional object passwordEncryptionProvider</param>
        /// <param name="passwordEncryptionAlgorithm">optional object passwordEncryptionAlgorithm</param>
        /// <param name="passwordEncryptionKeyLength">optional object passwordEncryptionKeyLength</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void SetPasswordEncryptionOptions(object passwordEncryptionProvider, object passwordEncryptionAlgorithm, object passwordEncryptionKeyLength)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetPasswordEncryptionOptions", passwordEncryptionProvider, passwordEncryptionAlgorithm, passwordEncryptionKeyLength);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual void RecheckSmartTags()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RecheckSmartTags");
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840370.aspx </remarks>
        /// <param name="recipients">optional object recipients</param>
        /// <param name="subject">optional object subject</param>
        /// <param name="showMessage">optional object showMessage</param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual void SendFaxOverInternet(object recipients, object subject, object showMessage)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SendFaxOverInternet", recipients, subject, showMessage);
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840370.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual void SendFaxOverInternet()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SendFaxOverInternet");
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840370.aspx </remarks>
        /// <param name="recipients">optional object recipients</param>
        [CustomMethod]
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual void SendFaxOverInternet(object recipients)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SendFaxOverInternet", recipients);
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840370.aspx </remarks>
        /// <param name="recipients">optional object recipients</param>
        /// <param name="subject">optional object subject</param>
        [CustomMethod]
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual void SendFaxOverInternet(object recipients, object subject)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SendFaxOverInternet", recipients, subject);
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836452.aspx </remarks>
        /// <param name="url">string url</param>
        /// <param name="importMap">NetOffice.ExcelApi.XmlMap importMap</param>
        /// <param name="overwrite">optional object overwrite</param>
        /// <param name="destination">optional object destination</param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImport(string url, out NetOffice.ExcelApi.XmlMap importMap, object overwrite, object destination)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true, false, false);
            object[] paramsArray = new object[] { url, new object(), overwrite, destination };

            NetOffice.ExcelApi.Enums.XlXmlImportResult returnItem = 
                InvokerService.InvokeInternal.ExecuteEnumMethodGetExtended<NetOffice.ExcelApi.Enums.XlXmlImportResult>(this, "XmlImport", paramsArray, modifiers);

            if (paramsArray[1] is MarshalByRefObject)
                importMap = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.XmlMap>(this, paramsArray[1], typeof(NetOffice.ExcelApi.XmlMap));
            else
                importMap = null;

            return returnItem;
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836452.aspx </remarks>
        /// <param name="url">string url</param>
        /// <param name="importMap">NetOffice.ExcelApi.XmlMap importMap</param>
        [CustomMethod]
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImport(string url, out NetOffice.ExcelApi.XmlMap importMap)
        {            
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true);
            object[] paramsArray = new object[] { url, new object() };

            NetOffice.ExcelApi.Enums.XlXmlImportResult returnItem = 
                InvokerService.InvokeInternal.ExecuteEnumMethodGetExtended<NetOffice.ExcelApi.Enums.XlXmlImportResult>(this, "XmlImport", paramsArray, modifiers);

            if (paramsArray[1] is MarshalByRefObject)
                importMap = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.XmlMap>(this, paramsArray[1], typeof(NetOffice.ExcelApi.XmlMap));
            else
                importMap = null;

            return returnItem;
         }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836452.aspx </remarks>
        /// <param name="url">string url</param>
        /// <param name="importMap">NetOffice.ExcelApi.XmlMap importMap</param>
        /// <param name="overwrite">optional object overwrite</param>
        [CustomMethod]
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImport(string url, out NetOffice.ExcelApi.XmlMap importMap, object overwrite)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true, false);
            object[] paramsArray = new object[] { url, new object(), overwrite };

            NetOffice.ExcelApi.Enums.XlXmlImportResult returnItem = 
                InvokerService.InvokeInternal.ExecuteEnumMethodGetExtended<NetOffice.ExcelApi.Enums.XlXmlImportResult>(this, "XmlImport", paramsArray, modifiers);

            if (paramsArray[1] is MarshalByRefObject)
                importMap = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.XmlMap>(this, paramsArray[1], typeof(NetOffice.ExcelApi.XmlMap));
            else
                importMap = null;

            return returnItem;
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837991.aspx </remarks>
        /// <param name="data">string data</param>
        /// <param name="importMap">NetOffice.ExcelApi.XmlMap importMap</param>
        /// <param name="overwrite">optional object overwrite</param>
        /// <param name="destination">optional object destination</param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImportXml(string data, out NetOffice.ExcelApi.XmlMap importMap, object overwrite, object destination)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true, false, false);
            object[] paramsArray = new object[] { data, new object(), overwrite, destination };

            NetOffice.ExcelApi.Enums.XlXmlImportResult returnItem = 
                InvokerService.InvokeInternal.ExecuteEnumMethodGetExtended<NetOffice.ExcelApi.Enums.XlXmlImportResult>(this, "XmlImportXml", paramsArray, modifiers);

            if (paramsArray[1] is MarshalByRefObject)
                importMap = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.XmlMap>(this, paramsArray[1], typeof(NetOffice.ExcelApi.XmlMap));
            else
                importMap = null;

            return returnItem;
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837991.aspx </remarks>
        /// <param name="data">string data</param>
        /// <param name="importMap">NetOffice.ExcelApi.XmlMap importMap</param>
        [CustomMethod]
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImportXml(string data, out NetOffice.ExcelApi.XmlMap importMap)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true);
            object[] paramsArray = new object[] { data, new object() };

            NetOffice.ExcelApi.Enums.XlXmlImportResult returnItem = 
                InvokerService.InvokeInternal.ExecuteEnumMethodGetExtended<NetOffice.ExcelApi.Enums.XlXmlImportResult>(this, "XmlImportXml", paramsArray, modifiers);

            if (paramsArray[1] is MarshalByRefObject)
                importMap = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.XmlMap>(this, paramsArray[1], typeof(NetOffice.ExcelApi.XmlMap));
            else
                importMap = null;

            return returnItem;
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837991.aspx </remarks>
        /// <param name="data">string data</param>
        /// <param name="importMap">NetOffice.ExcelApi.XmlMap importMap</param>
        /// <param name="overwrite">optional object overwrite</param>
        [CustomMethod]
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlXmlImportResult XmlImportXml(string data, out NetOffice.ExcelApi.XmlMap importMap, object overwrite)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true, false);
            object[] paramsArray = new object[] { data, new object(), overwrite };

            NetOffice.ExcelApi.Enums.XlXmlImportResult returnItem = 
                InvokerService.InvokeInternal.ExecuteEnumMethodGetExtended<NetOffice.ExcelApi.Enums.XlXmlImportResult>(this, "XmlImportXml", paramsArray, modifiers);

            if (paramsArray[1] is MarshalByRefObject)
                importMap = Factory.CreateKnownObjectFromComProxy<NetOffice.ExcelApi.XmlMap>(this, paramsArray[1], typeof(NetOffice.ExcelApi.XmlMap));
            else
                importMap = null;

            return returnItem;
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834616.aspx </remarks>
        /// <param name="filename">string filename</param>
        /// <param name="map">NetOffice.ExcelApi.XmlMap map</param>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual void SaveAsXMLData(string filename, NetOffice.ExcelApi.XmlMap map)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAsXMLData", filename, map);
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196845.aspx </remarks>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public virtual void ToggleFormsDesign()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ToggleFormsDesign");
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
        public virtual void __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "__PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile, collate });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void __PrintOut()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "__PrintOut");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void __PrintOut(object from)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "__PrintOut", from);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void __PrintOut(object from, object to)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "__PrintOut", from, to);
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
        public virtual void __PrintOut(object from, object to, object copies)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "__PrintOut", from, to, copies);
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
        public virtual void __PrintOut(object from, object to, object copies, object preview)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "__PrintOut", from, to, copies, preview);
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
        public virtual void __PrintOut(object from, object to, object copies, object preview, object activePrinter)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "__PrintOut", new object[] { from, to, copies, preview, activePrinter });
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
        public virtual void __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "__PrintOut", new object[] { from, to, copies, preview, activePrinter, printToFile });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        /// <param name="sharingPassword">optional object sharingPassword</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void _ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup, object sharingPassword)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_ProtectSharing", new object[] { filename, password, writeResPassword, readOnlyRecommended, createBackup, sharingPassword });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void _ProtectSharing()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_ProtectSharing");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void _ProtectSharing(object filename)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_ProtectSharing", filename);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void _ProtectSharing(object filename, object password)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_ProtectSharing", filename, password);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void _ProtectSharing(object filename, object password, object writeResPassword)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_ProtectSharing", filename, password, writeResPassword);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void _ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_ProtectSharing", filename, password, writeResPassword, readOnlyRecommended);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">optional object filename</param>
        /// <param name="password">optional object password</param>
        /// <param name="writeResPassword">optional object writeResPassword</param>
        /// <param name="readOnlyRecommended">optional object readOnlyRecommended</param>
        /// <param name="createBackup">optional object createBackup</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void _ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended, object createBackup)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_ProtectSharing", new object[] { filename, password, writeResPassword, readOnlyRecommended, createBackup });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840327.aspx </remarks>
        /// <param name="removeDocInfoType">NetOffice.ExcelApi.Enums.XlRemoveDocInfoType removeDocInfoType</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void RemoveDocumentInformation(NetOffice.ExcelApi.Enums.XlRemoveDocInfoType removeDocInfoType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveDocumentInformation", removeDocInfoType);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="comments">optional object comments</param>
        /// <param name="makePublic">optional object makePublic</param>
        /// <param name="versionType">optional object versionType</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void CheckInWithVersion(object saveChanges, object comments, object makePublic, object versionType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckInWithVersion", saveChanges, comments, makePublic, versionType);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void CheckInWithVersion()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckInWithVersion");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void CheckInWithVersion(object saveChanges)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckInWithVersion", saveChanges);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="comments">optional object comments</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void CheckInWithVersion(object saveChanges, object comments)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckInWithVersion", saveChanges, comments);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196878.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="comments">optional object comments</param>
        /// <param name="makePublic">optional object makePublic</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void CheckInWithVersion(object saveChanges, object comments, object makePublic)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckInWithVersion", saveChanges, comments, makePublic);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838567.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void LockServerFile()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "LockServerFile");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835507.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.WorkflowTasks GetWorkflowTasks()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.WorkflowTasks>(this, "GetWorkflowTasks", typeof(NetOffice.OfficeApi.WorkflowTasks));
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837818.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.WorkflowTemplates GetWorkflowTemplates()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.WorkflowTemplates>(this, "GetWorkflowTemplates", typeof(NetOffice.OfficeApi.WorkflowTemplates));
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194014.aspx </remarks>
        /// <param name="filename">string filename</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void ApplyTheme(string filename)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyTheme", filename);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820742.aspx </remarks>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void EnableConnections()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "EnableConnections");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
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
        public virtual void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish, object fixedFormatExtClassPtr)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish, fixedFormatExtClassPtr });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", type);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", type, filename);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", type, filename, quality);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", type, filename, quality, includeDocProperties);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { type, filename, quality, includeDocProperties, ignorePrintAreas });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        /// <param name="from">optional object from</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { type, filename, quality, includeDocProperties, ignorePrintAreas, from });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFixedFormatType type</param>
        /// <param name="filename">optional object filename</param>
        /// <param name="quality">optional object quality</param>
        /// <param name="includeDocProperties">optional object includeDocProperties</param>
        /// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { type, filename, quality, includeDocProperties, ignorePrintAreas, from, to });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198122.aspx </remarks>
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
        public virtual void ExportAsFixedFormat(NetOffice.ExcelApi.Enums.XlFixedFormatType type, object filename, object quality, object includeDocProperties, object ignorePrintAreas, object from, object to, object openAfterPublish)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { type, filename, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish });
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual void Dummy26()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Dummy26");
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual void Dummy27()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Dummy27");
        }

        #endregion

        #pragma warning restore
    }
}