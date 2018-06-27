using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi.Behind
{
    /// <summary>
    /// IVDocument
    /// </summary>
    [SyntaxBypass]
    public class IVDocument_ : COMObject, NetOffice.VisioApi.IVDocument_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public IVDocument_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Double get_LeftMargin(object unitsNameOrCode)
        {
            return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "LeftMargin", unitsNameOrCode);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        /// <param name="value">optional double value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void set_LeftMargin(object unitsNameOrCode, Double value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "LeftMargin", unitsNameOrCode, value);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_LeftMargin
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_LeftMargin")]
        public Double LeftMargin(object unitsNameOrCode)
        {
            return get_LeftMargin(unitsNameOrCode);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Double get_RightMargin(object unitsNameOrCode)
        {
            return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "RightMargin", unitsNameOrCode);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        /// <param name="value">optional double value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void set_RightMargin(object unitsNameOrCode, Double value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "RightMargin", unitsNameOrCode, value);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_RightMargin
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_RightMargin")]
        public Double RightMargin(object unitsNameOrCode)
        {
            return get_RightMargin(unitsNameOrCode);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Double get_TopMargin(object unitsNameOrCode)
        {
            return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "TopMargin", unitsNameOrCode);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        /// <param name="value">optional double value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void set_TopMargin(object unitsNameOrCode, Double value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "TopMargin", unitsNameOrCode, value);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_TopMargin
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_TopMargin")]
        public Double TopMargin(object unitsNameOrCode)
        {
            return get_TopMargin(unitsNameOrCode);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Double get_BottomMargin(object unitsNameOrCode)
        {
            return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "BottomMargin", unitsNameOrCode);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        /// <param name="value">optional double value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void set_BottomMargin(object unitsNameOrCode, Double value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "BottomMargin", unitsNameOrCode, value);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_BottomMargin
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_BottomMargin")]
        public Double BottomMargin(object unitsNameOrCode)
        {
            return get_BottomMargin(unitsNameOrCode);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="bstrPassword">optional object bstrPassword</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.VisioApi.Enums.VisProtection get_Protection(object bstrPassword)
        {
            return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisProtection>(this, "Protection", bstrPassword);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="bstrPassword">optional object bstrPassword</param>
        /// <param name="value">optional VisProtection value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void set_Protection(object bstrPassword, NetOffice.VisioApi.Enums.VisProtection value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "Protection", bstrPassword, value);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_Protection
        /// </summary>
        /// <param name="bstrPassword">optional object bstrPassword</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_Protection")]
        public NetOffice.VisioApi.Enums.VisProtection Protection(object bstrPassword)
        {
            return get_Protection(bstrPassword);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Double get_HeaderMargin(object unitsNameOrCode)
        {
            return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "HeaderMargin", unitsNameOrCode);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        /// <param name="value">optional double value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void set_HeaderMargin(object unitsNameOrCode, Double value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "HeaderMargin", unitsNameOrCode, value);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_HeaderMargin
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_HeaderMargin")]
        public Double HeaderMargin(object unitsNameOrCode)
        {
            return get_HeaderMargin(unitsNameOrCode);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Double get_FooterMargin(object unitsNameOrCode)
        {
            return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "FooterMargin", unitsNameOrCode);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        /// <param name="value">optional double value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void set_FooterMargin(object unitsNameOrCode, Double value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "FooterMargin", unitsNameOrCode, value);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_FooterMargin
        /// </summary>
        /// <param name="unitsNameOrCode">optional object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_FooterMargin")]
        public Double FooterMargin(object unitsNameOrCode)
        {
            return get_FooterMargin(unitsNameOrCode);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="bstrExistingPassword">optional object bstrExistingPassword</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public string get_Password(object bstrExistingPassword)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Password", bstrExistingPassword);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="bstrExistingPassword">optional object bstrExistingPassword</param>
        /// <param name="value">optional string value</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void set_Password(object bstrExistingPassword, string value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "Password", bstrExistingPassword, value);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_Password
        /// </summary>
        /// <param name="bstrExistingPassword">optional object bstrExistingPassword</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_Password")]
        public string Password(object bstrExistingPassword)
        {
            return get_Password(bstrExistingPassword);
        }

        #endregion
    }

    /// <summary>
    /// DispatchInterface IVDocument 
    /// SupportByVersion Visio, 11,12,14,15,16
    /// </summary>
    [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
    public class IVDocument : IVDocument_, NetOffice.VisioApi.IVDocument
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
                    _contractType = typeof(NetOffice.VisioApi.IVDocument);
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
                    _type = typeof(IVDocument);
                return _type;
            }
        }

        #endregion

        #region Ctor
      
        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public IVDocument() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVApplication Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 Stat
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Stat");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 ObjectType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ObjectType");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 InPlace
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "InPlace");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVMasters Masters
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVMasters>(this, "Masters");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVPages Pages
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVPages>(this, "Pages");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVStyles Styles
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVStyles>(this, "Styles");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Name
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Path
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Path");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string FullName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FullName");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 Index
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Index");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 old_Saved
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "old_Saved");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "old_Saved", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 ReadOnly
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ReadOnly");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int32 old_Version
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "old_Version");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "old_Version", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Title
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
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Subject
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
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Creator");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Creator", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Keywords
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
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Description
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Description");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Description", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVUIObject CustomMenus
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVUIObject>(this, "CustomMenus");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string CustomMenusFile
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CustomMenusFile");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CustomMenusFile", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVUIObject CustomToolbars
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVUIObject>(this, "CustomToolbars");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string CustomToolbarsFile
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CustomToolbarsFile");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CustomToolbarsFile", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVFonts Fonts
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVFonts>(this, "Fonts");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVColors Colors
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVColors>(this, "Colors");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVEventList EventList
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVEventList>(this, "EventList");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Template
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Template");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 old_SavePreviewMode
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "old_SavePreviewMode");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "old_SavePreviewMode", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Double LeftMargin
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "LeftMargin");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LeftMargin", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Double RightMargin
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "RightMargin");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RightMargin", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Double TopMargin
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "TopMargin");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TopMargin", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Double BottomMargin
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "BottomMargin");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BottomMargin", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 old_PrintLandscape
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "old_PrintLandscape");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "old_PrintLandscape", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 old_PrintCenteredH
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "old_PrintCenteredH");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "old_PrintCenteredH", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 old_PrintCenteredV
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "old_PrintCenteredV");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "old_PrintCenteredV", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Double PrintScale
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "PrintScale");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintScale", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 old_PrintFitOnPages
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "old_PrintFitOnPages");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "old_PrintFitOnPages", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 PrintPagesAcross
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "PrintPagesAcross");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintPagesAcross", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 PrintPagesDown
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "PrintPagesDown");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintPagesDown", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string DefaultStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string DefaultLineStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultLineStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultLineStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string DefaultFillStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultFillStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultFillStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string DefaultTextStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultTextStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultTextStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 PersistsEvents
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "PersistsEvents");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        public object VBProject
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "VBProject");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="unitsNameOrCode">object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Double get_PaperWidth(object unitsNameOrCode)
        {
            return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "PaperWidth", unitsNameOrCode);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_PaperWidth
        /// </summary>
        /// <param name="unitsNameOrCode">object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_PaperWidth")]
        public Double PaperWidth(object unitsNameOrCode)
        {
            return get_PaperWidth(unitsNameOrCode);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="unitsNameOrCode">object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Double get_PaperHeight(object unitsNameOrCode)
        {
            return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "PaperHeight", unitsNameOrCode);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_PaperHeight
        /// </summary>
        /// <param name="unitsNameOrCode">object unitsNameOrCode</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_PaperHeight")]
        public Double PaperHeight(object unitsNameOrCode)
        {
            return get_PaperHeight(unitsNameOrCode);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 old_PaperSize
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "old_PaperSize");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "old_PaperSize", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public string CodeName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CodeName");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int16 old_Mode
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "old_Mode");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "old_Mode", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVOLEObjects OLEObjects
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVOLEObjects>(this, "OLEObjects");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Manager
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Manager");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Manager", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Company
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Company");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Company", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Category
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Category");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Category", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string HyperlinkBase
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HyperlinkBase");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HyperlinkBase", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape DocumentSheet
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "DocumentSheet");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        public object Container
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Container");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string ClassID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ClassID");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string ProgID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ProgID");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVMasterShortcuts MasterShortcuts
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVMasterShortcuts>(this, "MasterShortcuts");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string AlternateNames
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AlternateNames");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AlternateNames", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVShape GestureFormatSheet
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "GestureFormatSheet");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool AutoRecover
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoRecover");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoRecover", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool Saved
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
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.Enums.VisDocVersions Version
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisDocVersions>(this, "Version");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Version", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.Enums.VisSavePreviewMode SavePreviewMode
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisSavePreviewMode>(this, "SavePreviewMode");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SavePreviewMode", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool PrintLandscape
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PrintLandscape");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintLandscape", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool PrintCenteredH
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PrintCenteredH");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintCenteredH", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool PrintCenteredV
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PrintCenteredV");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintCenteredV", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool PrintFitOnPages
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PrintFitOnPages");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintFitOnPages", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.Enums.VisPaperSizes PaperSize
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisPaperSizes>(this, "PaperSize");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PaperSize", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.Enums.VisDocModeArgs Mode
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisDocModeArgs>(this, "Mode");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Mode", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool SnapEnabled
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SnapEnabled");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SnapEnabled", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.Enums.VisSnapSettings SnapSettings
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisSnapSettings>(this, "SnapSettings");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SnapSettings", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.Enums.VisSnapExtensions SnapExtensions
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisSnapExtensions>(this, "SnapExtensions");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SnapExtensions", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Double[] SnapAngles
        {
            get
            {
                object[] paramsArray = null;
                object returnItem = (object)Invoker.PropertyGet(this, "SnapAngles", paramsArray);
                return (Double[])returnItem;
            }
            set
            {
                object[] paramsArray = Invoker.ValidateParamsArray(value);
                Invoker.PropertySet(this, "SnapAngles", paramsArray);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool GlueEnabled
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "GlueEnabled");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GlueEnabled", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.Enums.VisGlueSettings GlueSettings
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisGlueSettings>(this, "GlueSettings");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "GlueSettings", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool DynamicGridEnabled
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DynamicGridEnabled");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DynamicGridEnabled", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string DefaultGuideStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultGuideStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultGuideStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.Enums.VisProtection Protection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisProtection>(this, "Protection");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Protection", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string Printer
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Printer");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Printer", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Int32 PrintCopies
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PrintCopies");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintCopies", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string HeaderLeft
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HeaderLeft");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HeaderLeft", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string HeaderCenter
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HeaderCenter");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HeaderCenter", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string HeaderRight
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HeaderRight");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HeaderRight", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Double HeaderMargin
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "HeaderMargin");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HeaderMargin", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string FooterLeft
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FooterLeft");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FooterLeft", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string FooterCenter
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FooterCenter");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FooterCenter", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public string FooterRight
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FooterRight");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FooterRight", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Double FooterMargin
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "FooterMargin");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FooterMargin", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), NativeResult]
        public stdole.Font HeaderFooterFont
        {
            get
            {
                object[] paramsArray = null;
                object returnItem = Invoker.PropertyGet(this, "HeaderFooterFont", paramsArray);
                return returnItem as stdole.Font;
            }
            set
            {
                object[] paramsArray = Invoker.ValidateParamsArray(value);
                Invoker.PropertySet(this, "HeaderFooterFont", paramsArray);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 HeaderFooterColor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "HeaderFooterColor");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HeaderFooterColor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public string Password
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
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), NativeResult]
        public stdole.Picture PreviewPicture
        {
            get
            {
                object[] paramsArray = null;
                object returnItem = Invoker.PropertyGet(this, "PreviewPicture", paramsArray);
                return returnItem as stdole.Picture;
            }
            set
            {
                object[] paramsArray = Invoker.ValidateParamsArray(value);
                Invoker.PropertySet(this, "PreviewPicture", paramsArray);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 BuildNumberCreated
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BuildNumberCreated");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 BuildNumberEdited
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BuildNumberEdited");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public DateTime TimeCreated
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDateTimePropertyGet(this, "TimeCreated");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public DateTime Time
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDateTimePropertyGet(this, "Time");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public DateTime TimeEdited
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDateTimePropertyGet(this, "TimeEdited");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public DateTime TimePrinted
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDateTimePropertyGet(this, "TimePrinted");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public DateTime TimeSaved
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDateTimePropertyGet(this, "TimeSaved");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool ContainsWorkspace
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ContainsWorkspace");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public object[] EmailRoutingData
        {
            get
            {
                object[] paramsArray = null;
                object returnItem = Invoker.PropertyGet(this, "EmailRoutingData", paramsArray);
                if ((null != returnItem) && (returnItem is MarshalByRefObject))
                {
                    ICOMObject[] newObject = Factory.CreateObjectArrayFromComProxy(this, (object[])returnItem, false);
                    return newObject;
                }
                else
                {
                    return (object[])returnItem;
                }
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public byte[] VBProjectData
        {
            get
            {
                object[] paramsArray = null;
                object returnItem = (object)Invoker.PropertyGet(this, "VBProjectData", paramsArray);
                return (byte[])returnItem;
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 SolutionXMLElementCount
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SolutionXMLElementCount");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public string get_SolutionXMLElementName(Int32 index)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SolutionXMLElementName", index);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_SolutionXMLElementName
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_SolutionXMLElementName")]
        public string SolutionXMLElementName(Int32 index)
        {
            return get_SolutionXMLElementName(index);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="elementName">string elementName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool get_SolutionXMLElementExists(string elementName)
        {
            return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SolutionXMLElementExists", elementName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_SolutionXMLElementExists
        /// </summary>
        /// <param name="elementName">string elementName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_SolutionXMLElementExists")]
        public bool SolutionXMLElementExists(string elementName)
        {
            return get_SolutionXMLElementExists(elementName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="elementName">string elementName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public string get_SolutionXMLElement(string elementName)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SolutionXMLElement", elementName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="elementName">string elementName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void set_SolutionXMLElement(string elementName, string value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "SolutionXMLElement", elementName, value);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Alias for get_SolutionXMLElement
        /// </summary>
        /// <param name="elementName">string elementName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), Redirect("get_SolutionXMLElement")]
        public string SolutionXMLElement(string elementName)
        {
            return get_SolutionXMLElement(elementName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 FullBuildNumberCreated
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "FullBuildNumberCreated");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 FullBuildNumberEdited
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "FullBuildNumberEdited");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 ID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ID");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool MacrosEnabled
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MacrosEnabled");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.Enums.VisZoomBehavior ZoomBehavior
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisZoomBehavior>(this, "ZoomBehavior");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ZoomBehavior", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public NetOffice.VisioApi.Enums.VisDocumentTypes Type
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VisioApi.Enums.VisDocumentTypes>(this, "Type");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 Language
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Language");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Language", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool RemovePersonalInformation
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
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool UndoEnabled
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UndoEnabled");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UndoEnabled", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        public object SharedWorkspace
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "SharedWorkspace");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16), ProxyResult]
        public object Sync
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Sync");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVDataRecordsets DataRecordsets
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDataRecordsets>(this, "DataRecordsets");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public bool ContainsWorkspaceEx
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ContainsWorkspaceEx");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ContainsWorkspaceEx", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public string DefaultSavePath
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultSavePath");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultSavePath", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        public string CustomUI
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CustomUI");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CustomUI", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        public string UserCustomUI
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "UserCustomUI");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UserCustomUI", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVServerPublishOptions ServerPublishOptions
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVServerPublishOptions>(this, "ServerPublishOptions");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVValidation Validation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVValidation>(this, "Validation");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        public Int32 DiagramServicesEnabled
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DiagramServicesEnabled");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DiagramServicesEnabled", value);
            }
        }

        /// <summary>
        /// SupportByVersion Visio 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 15, 16)]
        public bool CompatibilityMode
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CompatibilityMode");
            }
        }

        /// <summary>
        /// SupportByVersion Visio 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Visio", 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVComments Comments
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVComments>(this, "Comments");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="objectToDrop">object objectToDrop</param>
        /// <param name="xPos">Int16 xPos</param>
        /// <param name="yPos">Int16 yPos</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVMaster Drop(object objectToDrop, Int16 xPos, Int16 yPos)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVMaster>(this, "Drop", objectToDrop, xPos, yPos);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 Save()
        {
            return InvokerService.InvokeInternal.ExecuteInt16MethodGet(this, "Save");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int16 SaveAs(string fileName)
        {
            return InvokerService.InvokeInternal.ExecuteInt16MethodGet(this, "SaveAs", fileName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Print()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Print");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Close()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Close");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="menusObject">NetOffice.VisioApi.IVUIObject menusObject</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void SetCustomMenus(NetOffice.VisioApi.IVUIObject menusObject)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetCustomMenus", menusObject);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void ClearCustomMenus()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearCustomMenus");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="toolbarsObject">NetOffice.VisioApi.IVUIObject toolbarsObject</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void SetCustomToolbars(NetOffice.VisioApi.IVUIObject toolbarsObject)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetCustomToolbars", toolbarsObject);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void ClearCustomToolbars()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearCustomToolbars");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="fileName">string fileName</param>
        /// <param name="saveFlags">Int16 saveFlags</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void SaveAsEx(string fileName, Int16 saveFlags)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAsEx", fileName, saveFlags);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="iD">Int16 iD</param>
        /// <param name="fileName">string fileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void GetIcon(Int16 iD, string fileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "GetIcon", iD, fileName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="iD">Int16 iD</param>
        /// <param name="index">Int16 index</param>
        /// <param name="fileName">string fileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void SetIcon(Int16 iD, Int16 index, string fileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetIcon", iD, index, fileName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        [BaseResult]
        public NetOffice.VisioApi.IVWindow OpenStencilWindow()
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "OpenStencilWindow");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="line">string line</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void ParseLine(string line)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ParseLine", line);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="line">string line</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void ExecuteLine(string line)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExecuteLine", line);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="target">string target</param>
        /// <param name="location">string location</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void FollowHyperlink45(string target, string location)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink45", target, location);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="address">string address</param>
        /// <param name="subAddress">string subAddress</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        /// <param name="frame">optional object frame</param>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="res1">optional object res1</param>
        /// <param name="res2">optional object res2</param>
        /// <param name="res3">optional object res3</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void FollowHyperlink(string address, string subAddress, object extraInfo, object frame, object newWindow, object res1, object res2, object res3)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", new object[] { address, subAddress, extraInfo, frame, newWindow, res1, res2, res3 });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="address">string address</param>
        /// <param name="subAddress">string subAddress</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void FollowHyperlink(string address, string subAddress)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address, subAddress);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="address">string address</param>
        /// <param name="subAddress">string subAddress</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void FollowHyperlink(string address, string subAddress, object extraInfo)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address, subAddress, extraInfo);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="address">string address</param>
        /// <param name="subAddress">string subAddress</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        /// <param name="frame">optional object frame</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void FollowHyperlink(string address, string subAddress, object extraInfo, object frame)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", address, subAddress, extraInfo, frame);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="address">string address</param>
        /// <param name="subAddress">string subAddress</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        /// <param name="frame">optional object frame</param>
        /// <param name="newWindow">optional object newWindow</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void FollowHyperlink(string address, string subAddress, object extraInfo, object frame, object newWindow)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", new object[] { address, subAddress, extraInfo, frame, newWindow });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="address">string address</param>
        /// <param name="subAddress">string subAddress</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        /// <param name="frame">optional object frame</param>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="res1">optional object res1</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void FollowHyperlink(string address, string subAddress, object extraInfo, object frame, object newWindow, object res1)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", new object[] { address, subAddress, extraInfo, frame, newWindow, res1 });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="address">string address</param>
        /// <param name="subAddress">string subAddress</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        /// <param name="frame">optional object frame</param>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="res1">optional object res1</param>
        /// <param name="res2">optional object res2</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void FollowHyperlink(string address, string subAddress, object extraInfo, object frame, object newWindow, object res1, object res2)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "FollowHyperlink", new object[] { address, subAddress, extraInfo, frame, newWindow, res1, res2 });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void ClearGestureFormatSheet()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearGestureFormatSheet");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nTargets">optional object nTargets</param>
        /// <param name="nActions">optional object nActions</param>
        /// <param name="nAlerts">optional object nAlerts</param>
        /// <param name="nFixes">optional object nFixes</param>
        /// <param name="bStopOnError">optional object bStopOnError</param>
        /// <param name="bLogFileName">optional object bLogFileName</param>
        /// <param name="nReserved">optional object nReserved</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Clean(object nTargets, object nActions, object nAlerts, object nFixes, object bStopOnError, object bLogFileName, object nReserved)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Clean", new object[] { nTargets, nActions, nAlerts, nFixes, bStopOnError, bLogFileName, nReserved });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Clean()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Clean");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nTargets">optional object nTargets</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Clean(object nTargets)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Clean", nTargets);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nTargets">optional object nTargets</param>
        /// <param name="nActions">optional object nActions</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Clean(object nTargets, object nActions)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Clean", nTargets, nActions);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nTargets">optional object nTargets</param>
        /// <param name="nActions">optional object nActions</param>
        /// <param name="nAlerts">optional object nAlerts</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Clean(object nTargets, object nActions, object nAlerts)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Clean", nTargets, nActions, nAlerts);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nTargets">optional object nTargets</param>
        /// <param name="nActions">optional object nActions</param>
        /// <param name="nAlerts">optional object nAlerts</param>
        /// <param name="nFixes">optional object nFixes</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Clean(object nTargets, object nActions, object nAlerts, object nFixes)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Clean", nTargets, nActions, nAlerts, nFixes);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nTargets">optional object nTargets</param>
        /// <param name="nActions">optional object nActions</param>
        /// <param name="nAlerts">optional object nAlerts</param>
        /// <param name="nFixes">optional object nFixes</param>
        /// <param name="bStopOnError">optional object bStopOnError</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Clean(object nTargets, object nActions, object nAlerts, object nFixes, object bStopOnError)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Clean", new object[] { nTargets, nActions, nAlerts, nFixes, bStopOnError });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nTargets">optional object nTargets</param>
        /// <param name="nActions">optional object nActions</param>
        /// <param name="nAlerts">optional object nAlerts</param>
        /// <param name="nFixes">optional object nFixes</param>
        /// <param name="bStopOnError">optional object bStopOnError</param>
        /// <param name="bLogFileName">optional object bLogFileName</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void Clean(object nTargets, object nActions, object nAlerts, object nFixes, object bStopOnError, object bLogFileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Clean", new object[] { nTargets, nActions, nAlerts, nFixes, bStopOnError, bLogFileName });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pSourceDoc">NetOffice.VisioApi.IVDocument pSourceDoc</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void CopyPreviewPicture(NetOffice.VisioApi.IVDocument pSourceDoc)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CopyPreviewPicture", pSourceDoc);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="elementName">string elementName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void DeleteSolutionXMLElement(string elementName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteSolutionXMLElement", elementName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public bool CanCheckIn()
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CanCheckIn");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="saveChanges">optional bool SaveChanges = true</param>
        /// <param name="comments">optional object comments</param>
        /// <param name="makePublic">optional bool MakePublic = false</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void CheckIn(object saveChanges, object comments, object makePublic)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn", saveChanges, comments, makePublic);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void CheckIn()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="saveChanges">optional bool SaveChanges = true</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void CheckIn(object saveChanges)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn", saveChanges);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="saveChanges">optional bool SaveChanges = true</param>
        /// <param name="comments">optional object comments</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void CheckIn(object saveChanges, object comments)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckIn", saveChanges, comments);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
        /// <param name="printerName">optional string PrinterName = </param>
        /// <param name="printToFile">optional bool PrintToFile = false</param>
        /// <param name="outputFileName">optional string OutputFileName = </param>
        /// <param name="copies">optional Int32 Copies = 1</param>
        /// <param name="collate">optional bool Collate = false</param>
        /// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper, object printerName, object printToFile, object outputFileName, object copies, object collate, object colorAsBlack)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { printRange, fromPage, toPage, scaleCurrentViewToPaper, printerName, printToFile, outputFileName, copies, collate, colorAsBlack });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", printRange);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", printRange, fromPage);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", printRange, fromPage, toPage);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", printRange, fromPage, toPage, scaleCurrentViewToPaper);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
        /// <param name="printerName">optional string PrinterName = </param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper, object printerName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { printRange, fromPage, toPage, scaleCurrentViewToPaper, printerName });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
        /// <param name="printerName">optional string PrinterName = </param>
        /// <param name="printToFile">optional bool PrintToFile = false</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper, object printerName, object printToFile)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { printRange, fromPage, toPage, scaleCurrentViewToPaper, printerName, printToFile });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
        /// <param name="printerName">optional string PrinterName = </param>
        /// <param name="printToFile">optional bool PrintToFile = false</param>
        /// <param name="outputFileName">optional string OutputFileName = </param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper, object printerName, object printToFile, object outputFileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { printRange, fromPage, toPage, scaleCurrentViewToPaper, printerName, printToFile, outputFileName });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
        /// <param name="printerName">optional string PrinterName = </param>
        /// <param name="printToFile">optional bool PrintToFile = false</param>
        /// <param name="outputFileName">optional string OutputFileName = </param>
        /// <param name="copies">optional Int32 Copies = 1</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper, object printerName, object printToFile, object outputFileName, object copies)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { printRange, fromPage, toPage, scaleCurrentViewToPaper, printerName, printToFile, outputFileName, copies });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
        /// <param name="printerName">optional string PrinterName = </param>
        /// <param name="printToFile">optional bool PrintToFile = false</param>
        /// <param name="outputFileName">optional string OutputFileName = </param>
        /// <param name="copies">optional Int32 Copies = 1</param>
        /// <param name="collate">optional bool Collate = false</param>
        [CustomMethod]
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper, object printerName, object printToFile, object outputFileName, object copies, object collate)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { printRange, fromPage, toPage, scaleCurrentViewToPaper, printerName, printToFile, outputFileName, copies, collate });
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrUndoScopeName">string bstrUndoScopeName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public Int32 BeginUndoScope(string bstrUndoScopeName)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "BeginUndoScope", bstrUndoScopeName);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="nScopeID">Int32 nScopeID</param>
        /// <param name="bCommit">bool bCommit</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void EndUndoScope(Int32 nScopeID, bool bCommit)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "EndUndoScope", nScopeID, bCommit);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="pUndoUnit">object pUndoUnit</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void AddUndoUnit(object pUndoUnit)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AddUndoUnit", pUndoUnit);
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void PurgeUndo()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PurgeUndo");
        }

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrScopeName">string bstrScopeName</param>
        [SupportByVersion("Visio", 11, 12, 14, 15, 16)]
        public void RenameCurrentScope(string bstrScopeName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RenameCurrentScope", bstrScopeName);
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="removeHiddenInfoItems">Int32 removeHiddenInfoItems</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void RemoveHiddenInformation(Int32 removeHiddenInfoItems)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RemoveHiddenInformation", removeHiddenInfoItems);
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="eType">NetOffice.VisioApi.Enums.VisThemeTypes eType</param>
        /// <param name="nameArray">String[] nameArray</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void GetThemeNames(NetOffice.VisioApi.Enums.VisThemeTypes eType, out String[] nameArray)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true);
            nameArray = null;
            object[] paramsArray = Invoker.ValidateParamsArray(eType, (object)nameArray);
            Invoker.Method(this, "GetThemeNames", paramsArray, modifiers);
            nameArray = (String[])paramsArray[1];
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="eType">NetOffice.VisioApi.Enums.VisThemeTypes eType</param>
        /// <param name="nameArray">String[] nameArray</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void GetThemeNamesU(NetOffice.VisioApi.Enums.VisThemeTypes eType, out String[] nameArray)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true);
            nameArray = null;
            object[] paramsArray = Invoker.ValidateParamsArray(eType, (object)nameArray);
            Invoker.Method(this, "GetThemeNamesU", paramsArray, modifiers);
            nameArray = (String[])paramsArray[1];
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public bool CanUndoCheckOut()
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CanUndoCheckOut");
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void UndoCheckOut()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UndoCheckOut");
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent intent</param>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
        /// <param name="includeBackground">optional bool IncludeBackground = true</param>
        /// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
        /// <param name="includeStructureTags">optional bool IncludeStructureTags = true</param>
        /// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
        /// <param name="fixedFormatExtClass">optional object fixedFormatExtClass</param>
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object colorAsBlack, object includeBackground, object includeDocumentProperties, object includeStructureTags, object useISO19005_1, object fixedFormatExtClass)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { fixedFormat, outputFileName, intent, printRange, fromPage, toPage, colorAsBlack, includeBackground, includeDocumentProperties, includeStructureTags, useISO19005_1, fixedFormatExtClass });
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent intent</param>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", fixedFormat, outputFileName, intent, printRange);
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent intent</param>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { fixedFormat, outputFileName, intent, printRange, fromPage });
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent intent</param>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { fixedFormat, outputFileName, intent, printRange, fromPage, toPage });
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent intent</param>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object colorAsBlack)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { fixedFormat, outputFileName, intent, printRange, fromPage, toPage, colorAsBlack });
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent intent</param>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
        /// <param name="includeBackground">optional bool IncludeBackground = true</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object colorAsBlack, object includeBackground)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { fixedFormat, outputFileName, intent, printRange, fromPage, toPage, colorAsBlack, includeBackground });
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent intent</param>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
        /// <param name="includeBackground">optional bool IncludeBackground = true</param>
        /// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object colorAsBlack, object includeBackground, object includeDocumentProperties)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { fixedFormat, outputFileName, intent, printRange, fromPage, toPage, colorAsBlack, includeBackground, includeDocumentProperties });
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent intent</param>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
        /// <param name="includeBackground">optional bool IncludeBackground = true</param>
        /// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
        /// <param name="includeStructureTags">optional bool IncludeStructureTags = true</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object colorAsBlack, object includeBackground, object includeDocumentProperties, object includeStructureTags)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { fixedFormat, outputFileName, intent, printRange, fromPage, toPage, colorAsBlack, includeBackground, includeDocumentProperties, includeStructureTags });
        }

        /// <summary>
        /// SupportByVersion Visio 12, 14, 15, 16
        /// </summary>
        /// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat</param>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent intent</param>
        /// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange printRange</param>
        /// <param name="fromPage">optional Int32 FromPage = 1</param>
        /// <param name="toPage">optional Int32 ToPage = -1</param>
        /// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
        /// <param name="includeBackground">optional bool IncludeBackground = true</param>
        /// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
        /// <param name="includeStructureTags">optional bool IncludeStructureTags = true</param>
        /// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
        [CustomMethod]
        [SupportByVersion("Visio", 12, 14, 15, 16)]
        public void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object colorAsBlack, object includeBackground, object includeDocumentProperties, object includeStructureTags, object useISO19005_1)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { fixedFormat, outputFileName, intent, printRange, fromPage, toPage, colorAsBlack, includeBackground, includeDocumentProperties, includeStructureTags, useISO19005_1 });
        }

        #endregion

        #pragma warning restore
    }
}
