using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface IFind 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class IFind : COMObject, NetOffice.OfficeApi.IFind
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
                    _type = typeof(IFind);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IFind() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public string SearchPath
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "SearchPath");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "SearchPath", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
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
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public bool SubDir
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "SubDir");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "SubDir", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public string Title
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Title");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Title", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public string Author
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Author");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Author", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public string Keywords
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Keywords");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Keywords", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public string Subject
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Subject");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Subject", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoFileFindOptions Options
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFileFindOptions>(this, "Options");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "Options", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public bool MatchCase
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "MatchCase");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "MatchCase", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public string Text
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Text");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Text", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public bool PatternMatch
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "PatternMatch");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "PatternMatch", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public object DateSavedFrom
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "DateSavedFrom");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "DateSavedFrom", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public object DateSavedTo
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "DateSavedTo");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "DateSavedTo", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public string SavedBy
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "SavedBy");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "SavedBy", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public object DateCreatedFrom
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "DateCreatedFrom");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "DateCreatedFrom", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public object DateCreatedTo
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "DateCreatedTo");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "DateCreatedTo", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoFileFindView View
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFileFindView>(this, "View");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "View", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoFileFindSortBy SortBy
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFileFindSortBy>(this, "SortBy");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "SortBy", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoFileFindListBy ListBy
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFileFindListBy>(this, "ListBy");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "ListBy", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 SelectedFile
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "SelectedFile");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "SelectedFile", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.IFoundFiles Results
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IFoundFiles>(this, "Results", typeof(NetOffice.OfficeApi.IFoundFiles));
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 FileType
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "FileType");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "FileType", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 Show()
        {
            return Factory.ExecuteInt32MethodGet(this, "Show");
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void Execute()
        {
            Factory.ExecuteMethod(this, "Execute");
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrQueryName">string bstrQueryName</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void Load(string bstrQueryName)
        {
            Factory.ExecuteMethod(this, "Load", bstrQueryName);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrQueryName">string bstrQueryName</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void Save(string bstrQueryName)
        {
            Factory.ExecuteMethod(this, "Save", bstrQueryName);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrQueryName">string bstrQueryName</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public void Delete(string bstrQueryName)
        {
            Factory.ExecuteMethod(this, "Delete", bstrQueryName);
        }

        #endregion

        #pragma warning restore
    }
}
