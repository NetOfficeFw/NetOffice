using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface _CustomTaskPane
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
    public class _CustomTaskPane : COMObject, NetOffice.OfficeApi._CustomTaskPane
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
                    _contractType = typeof(NetOffice.OfficeApi._CustomTaskPane);
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
                    _type = typeof(_CustomTaskPane);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _CustomTaskPane() : base()
		{

		}

        #endregion

        #region Events

        /// <summary>
        /// Occurs after Delete for the proxy has been called
        /// </summary>
        /// <remarks>The event occurs for the proxy instance only.</remarks>
        [CustomEvent]
        public event _CustomTaskPaneDeleteHandler AfterDelete;

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861137.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string Title
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Title");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862545.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        public virtual object Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Application");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862803.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        public virtual object Window
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Window");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865256.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Native COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861783.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), NativeResult]
        public virtual object ContentControl
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteObjectPropertyGet(this, "ContentControl");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860235.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 Height
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Height");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Height", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865362.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 Width
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Width");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Width", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861841.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoCTPDockPosition DockPosition
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoCTPDockPosition>(this, "DockPosition");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DockPosition", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861088.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoCTPDockPositionRestrict DockPositionRestrict
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoCTPDockPositionRestrict>(this, "DockPositionRestrict");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DockPositionRestrict", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862399.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Delete()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
            var handler = AfterDelete;
            if (null != handler)
                handler(this);
        }

        #endregion

        #pragma warning restore
    }
}
