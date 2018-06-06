using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi.Behind
{
    /// <summary>
    /// DispatchInterface _VBComponent_Old
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
    public class _VBComponent_Old : NetOffice.VBIDEApi.Behind._VBComponents_Old, NetOffice.VBIDEApi._VBComponent_Old
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
                    _type = typeof(_VBComponent_Old);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _VBComponent_Old() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual bool Saved
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "Saved");
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual string Name
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
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3), ProxyResult]
        public virtual object Designer
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Designer");
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual NetOffice.VBIDEApi.CodeModule CodeModule
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.CodeModule>(this, "CodeModule", typeof(NetOffice.VBIDEApi.CodeModule));
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual NetOffice.VBIDEApi.Enums.vbext_ComponentType Type
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.VBIDEApi.Enums.vbext_ComponentType>(this, "Type");
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual NetOffice.VBIDEApi.VBE VBE
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
        public virtual NetOffice.VBIDEApi.VBComponents Collection
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBComponents>(this, "Collection", typeof(NetOffice.VBIDEApi.VBComponents));
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual bool HasOpenDesigner
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "HasOpenDesigner");
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual NetOffice.VBIDEApi.Properties Properties
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.Properties>(this, "Properties", typeof(NetOffice.VBIDEApi.Properties));
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual void Export(string fileName)
        {
            Factory.ExecuteMethod(this, "Export", fileName);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual NetOffice.VBIDEApi.Window DesignerWindow()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.VBIDEApi.Window>(this, "DesignerWindow", typeof(NetOffice.VBIDEApi.Window));
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual void Activate()
        {
            Factory.ExecuteMethod(this, "Activate");
        }

        #endregion

        #pragma warning restore
    }
}
