using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi.Behind
{
    /// <summary>
    /// DispatchInterface _VBProject_Old
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
    public class _VBProject_Old :  NetOffice.VBIDEApi.Behind._ProjectTemplate, NetOffice.VBIDEApi._VBProject_Old
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
                    _type = typeof(_VBProject_Old);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _VBProject_Old() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public string HelpFile
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "HelpFile");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HelpFile", value);
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public Int32 HelpContextID
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "HelpContextID");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HelpContextID", value);
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public string Description
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Description");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Description", value);
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public NetOffice.VBIDEApi.Enums.vbext_VBAMode Mode
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.VBIDEApi.Enums.vbext_VBAMode>(this, "Mode");
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public NetOffice.VBIDEApi.References References
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.References>(this, "References", typeof(NetOffice.VBIDEApi.References));
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public string Name
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Name");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Name", value);
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public NetOffice.VBIDEApi.VBE VBE
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBE>(this, "VBE", typeof(NetOffice.VBIDEApi.VBE));
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public NetOffice.VBIDEApi.VBProjects Collection
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBProjects>(this, "Collection", typeof(NetOffice.VBIDEApi.VBProjects));
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public NetOffice.VBIDEApi.Enums.vbext_ProjectProtection Protection
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.VBIDEApi.Enums.vbext_ProjectProtection>(this, "Protection");
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public bool Saved
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "Saved");
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public NetOffice.VBIDEApi.VBComponents VBComponents
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBComponents>(this, "VBComponents", typeof(NetOffice.VBIDEApi.VBComponents));
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}
