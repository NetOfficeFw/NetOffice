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
                return Factory.ExecuteBoolPropertyGet(this, "SearchSubFolders");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "SearchSubFolders", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "MatchTextExactly");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "MatchTextExactly", value);
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
                return Factory.ExecuteBoolPropertyGet(this, "MatchAllWordForms");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "MatchAllWordForms", value);
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
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoFileType FileType
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFileType>(this, "FileType");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "FileType", value);
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
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoLastModified>(this, "LastModified");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "LastModified", value);
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
                return Factory.ExecuteStringPropertyGet(this, "TextOrProperty");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "TextOrProperty", value);
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
                return Factory.ExecuteStringPropertyGet(this, "LookIn");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "LookIn", value);
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
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FoundFiles>(this, "FoundFiles", typeof(NetOffice.OfficeApi.FoundFiles));
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
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PropertyTests>(this, "PropertyTests", typeof(NetOffice.OfficeApi.PropertyTests));
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
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SearchScopes>(this, "SearchScopes", typeof(NetOffice.OfficeApi.SearchScopes));
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
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SearchFolders>(this, "SearchFolders", typeof(NetOffice.OfficeApi.SearchFolders));
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
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileTypes>(this, "FileTypes", typeof(NetOffice.OfficeApi.FileTypes));
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
            return Factory.ExecuteInt32MethodGet(this, "Execute", sortBy, sortOrder, alwaysAccurate);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Execute()
        {
            return Factory.ExecuteInt32MethodGet(this, "Execute");
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="sortBy">optional NetOffice.OfficeApi.Enums.MsoSortBy SortBy = 1</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Execute(object sortBy)
        {
            return Factory.ExecuteInt32MethodGet(this, "Execute", sortBy);
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
            return Factory.ExecuteInt32MethodGet(this, "Execute", sortBy, sortOrder);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void NewSearch()
        {
            Factory.ExecuteMethod(this, "NewSearch");
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual void RefreshScopes()
        {
            Factory.ExecuteMethod(this, "RefreshScopes");
        }

        #endregion

        #pragma warning restore
    }
}
