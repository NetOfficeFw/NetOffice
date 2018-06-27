using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface FileSearch 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class FileSearch : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.FileSearch
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
                    _contractType = typeof(NetOffice.OfficeApi.FileSearch);
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
                    _type = typeof(FileSearch);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public FileSearch() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool SearchSubFolders
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SearchSubFolders");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SearchSubFolders", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool MatchTextExactly
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MatchTextExactly");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MatchTextExactly", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool MatchAllWordForms
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MatchAllWordForms");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MatchAllWordForms", value);
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
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoFileType FileType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFileType>(this, "FileType");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "FileType", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoLastModified LastModified
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoLastModified>(this, "LastModified");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "LastModified", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string TextOrProperty
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "TextOrProperty");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TextOrProperty", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string LookIn
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "LookIn");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LookIn", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.FoundFiles FoundFiles
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FoundFiles>(this, "FoundFiles", typeof(NetOffice.OfficeApi.FoundFiles));
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.PropertyTests PropertyTests
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PropertyTests>(this, "PropertyTests", typeof(NetOffice.OfficeApi.PropertyTests));
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.SearchScopes SearchScopes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SearchScopes>(this, "SearchScopes", typeof(NetOffice.OfficeApi.SearchScopes));
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.SearchFolders SearchFolders
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SearchFolders>(this, "SearchFolders", typeof(NetOffice.OfficeApi.SearchFolders));
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.FileTypes FileTypes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileTypes>(this, "FileTypes", typeof(NetOffice.OfficeApi.FileTypes));
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sortBy">optional NetOffice.OfficeApi.Enums.MsoSortBy SortBy = 1</param>
        /// <param name="sortOrder">optional NetOffice.OfficeApi.Enums.MsoSortOrder SortOrder = 1</param>
        /// <param name="alwaysAccurate">optional bool AlwaysAccurate = true</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Execute(object sortBy, object sortOrder, object alwaysAccurate)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Execute", sortBy, sortOrder, alwaysAccurate);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Execute()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Execute");
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sortBy">optional NetOffice.OfficeApi.Enums.MsoSortBy SortBy = 1</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Execute(object sortBy)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Execute", sortBy);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sortBy">optional NetOffice.OfficeApi.Enums.MsoSortBy SortBy = 1</param>
        /// <param name="sortOrder">optional NetOffice.OfficeApi.Enums.MsoSortOrder SortOrder = 1</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Execute(object sortBy, object sortOrder)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Execute", sortBy, sortOrder);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void NewSearch()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "NewSearch");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void RefreshScopes()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RefreshScopes");
        }

        #endregion

        #pragma warning restore
    }
}
