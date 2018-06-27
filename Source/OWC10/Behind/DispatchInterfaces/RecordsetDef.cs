using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api.Behind
{
    /// <summary>
    /// RecordsetDef
    /// </summary>
    [SyntaxBypass]
    public class RecordsetDef_ : COMObject, NetOffice.OWC10Api.RecordsetDef_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public RecordsetDef_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        /// <param name="fetchType">optional NetOffice.OWC10Api.Enums.DscFetchTypeEnum fetchType</param>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_ShapeText(object fetchType)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ShapeText", fetchType);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_ShapeText
        /// </summary>
        /// <param name="fetchType">optional NetOffice.OWC10Api.Enums.DscFetchTypeEnum fetchType</param>
        [SupportByVersion("OWC10", 1), Redirect("get_ShapeText")]
        public virtual string ShapeText(object fetchType)
        {
            return get_ShapeText(fetchType);
        }

        #endregion
    }

    /// <summary>
    /// DispatchInterface RecordsetDef 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class RecordsetDef : RecordsetDef_, NetOffice.OWC10Api.RecordsetDef
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
                    _contractType = typeof(NetOffice.OWC10Api.RecordsetDef);
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
                    _type = typeof(RecordsetDef);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public RecordsetDef() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
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
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string ShapeText
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ShapeText");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual string CommandText
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CommandText");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool DataEntry
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DataEntry");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DataEntry", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual string ServerFilter
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ServerFilter");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ServerFilter", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.PageRowsource PrimaryPageRowsource
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PageRowsource>(this, "PrimaryPageRowsource", typeof(NetOffice.OWC10Api.PageRowsource));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.SublistRelationships SublistRelationships
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.SublistRelationships>(this, "SublistRelationships", typeof(NetOffice.OWC10Api.SublistRelationships));            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.PageFields PageFields
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PageFields>(this, "PageFields", typeof(NetOffice.OWC10Api.PageFields));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.RecordsetDef ParentRecordsetDef
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.RecordsetDef>(this, "ParentRecordsetDef", typeof(NetOffice.OWC10Api.RecordsetDef));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.GroupingDefs GroupingDefs
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.GroupingDefs>(this, "GroupingDefs", typeof(NetOffice.OWC10Api.GroupingDefs));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.ParameterValues ParameterValues
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ParameterValues>(this, "ParameterValues", typeof(NetOffice.OWC10Api.ParameterValues));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.PageRowsources PageRowsources
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PageRowsources>(this, "PageRowsources", typeof(NetOffice.OWC10Api.PageRowsources));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual string UniqueTable
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "UniqueTable");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UniqueTable", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual string ResyncCommand
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ResyncCommand");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ResyncCommand", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.RecordsetDef Demote()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.RecordsetDef>(this, "Demote", typeof(NetOffice.OWC10Api.RecordsetDef));
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void Delete()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
        }

        #endregion

        #pragma warning restore
    }
}

