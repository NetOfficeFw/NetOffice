using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface TablesOfAuthorities 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837712.aspx </remarks>
	public class TablesOfAuthorities : COMObject, NetOffice.WordApi.TablesOfAuthorities
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
                    _contractType = typeof(NetOffice.WordApi.TablesOfAuthorities);
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
                    _type = typeof(TablesOfAuthorities);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public TablesOfAuthorities() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820743.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845059.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838690.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837691.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839360.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdToaFormat Format
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdToaFormat>(this, "Format");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Format", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.WordApi.TableOfAuthorities this[Int32 index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfAuthorities>(this, "Item", typeof(NetOffice.WordApi.TableOfAuthorities), index);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		/// <param name="bookmark">optional object bookmark</param>
		/// <param name="passim">optional object passim</param>
		/// <param name="keepEntryFormatting">optional object keepEntryFormatting</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="includeSequenceName">optional object includeSequenceName</param>
		/// <param name="entrySeparator">optional object entrySeparator</param>
		/// <param name="pageRangeSeparator">optional object pageRangeSeparator</param>
		/// <param name="includeCategoryHeader">optional object includeCategoryHeader</param>
		/// <param name="pageNumberSeparator">optional object pageNumberSeparator</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting, object separator, object includeSequenceName, object entrySeparator, object pageRangeSeparator, object includeCategoryHeader, object pageNumberSeparator)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfAuthorities>(this, "Add", typeof(NetOffice.WordApi.TableOfAuthorities), new object[]{ range, category, bookmark, passim, keepEntryFormatting, separator, includeSequenceName, entrySeparator, pageRangeSeparator, includeCategoryHeader, pageNumberSeparator });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfAuthorities>(this, "Add", typeof(NetOffice.WordApi.TableOfAuthorities), range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfAuthorities>(this, "Add", typeof(NetOffice.WordApi.TableOfAuthorities), range, category);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		/// <param name="bookmark">optional object bookmark</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfAuthorities>(this, "Add", typeof(NetOffice.WordApi.TableOfAuthorities), range, category, bookmark);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		/// <param name="bookmark">optional object bookmark</param>
		/// <param name="passim">optional object passim</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfAuthorities>(this, "Add", typeof(NetOffice.WordApi.TableOfAuthorities), range, category, bookmark, passim);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		/// <param name="bookmark">optional object bookmark</param>
		/// <param name="passim">optional object passim</param>
		/// <param name="keepEntryFormatting">optional object keepEntryFormatting</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfAuthorities>(this, "Add", typeof(NetOffice.WordApi.TableOfAuthorities), new object[]{ range, category, bookmark, passim, keepEntryFormatting });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		/// <param name="bookmark">optional object bookmark</param>
		/// <param name="passim">optional object passim</param>
		/// <param name="keepEntryFormatting">optional object keepEntryFormatting</param>
		/// <param name="separator">optional object separator</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting, object separator)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfAuthorities>(this, "Add", typeof(NetOffice.WordApi.TableOfAuthorities), new object[]{ range, category, bookmark, passim, keepEntryFormatting, separator });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		/// <param name="bookmark">optional object bookmark</param>
		/// <param name="passim">optional object passim</param>
		/// <param name="keepEntryFormatting">optional object keepEntryFormatting</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="includeSequenceName">optional object includeSequenceName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting, object separator, object includeSequenceName)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfAuthorities>(this, "Add", typeof(NetOffice.WordApi.TableOfAuthorities), new object[]{ range, category, bookmark, passim, keepEntryFormatting, separator, includeSequenceName });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		/// <param name="bookmark">optional object bookmark</param>
		/// <param name="passim">optional object passim</param>
		/// <param name="keepEntryFormatting">optional object keepEntryFormatting</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="includeSequenceName">optional object includeSequenceName</param>
		/// <param name="entrySeparator">optional object entrySeparator</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting, object separator, object includeSequenceName, object entrySeparator)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfAuthorities>(this, "Add", typeof(NetOffice.WordApi.TableOfAuthorities), new object[]{ range, category, bookmark, passim, keepEntryFormatting, separator, includeSequenceName, entrySeparator });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		/// <param name="bookmark">optional object bookmark</param>
		/// <param name="passim">optional object passim</param>
		/// <param name="keepEntryFormatting">optional object keepEntryFormatting</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="includeSequenceName">optional object includeSequenceName</param>
		/// <param name="entrySeparator">optional object entrySeparator</param>
		/// <param name="pageRangeSeparator">optional object pageRangeSeparator</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting, object separator, object includeSequenceName, object entrySeparator, object pageRangeSeparator)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfAuthorities>(this, "Add", typeof(NetOffice.WordApi.TableOfAuthorities), new object[]{ range, category, bookmark, passim, keepEntryFormatting, separator, includeSequenceName, entrySeparator, pageRangeSeparator });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="category">optional object category</param>
		/// <param name="bookmark">optional object bookmark</param>
		/// <param name="passim">optional object passim</param>
		/// <param name="keepEntryFormatting">optional object keepEntryFormatting</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="includeSequenceName">optional object includeSequenceName</param>
		/// <param name="entrySeparator">optional object entrySeparator</param>
		/// <param name="pageRangeSeparator">optional object pageRangeSeparator</param>
		/// <param name="includeCategoryHeader">optional object includeCategoryHeader</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting, object separator, object includeSequenceName, object entrySeparator, object pageRangeSeparator, object includeCategoryHeader)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfAuthorities>(this, "Add", typeof(NetOffice.WordApi.TableOfAuthorities), new object[]{ range, category, bookmark, passim, keepEntryFormatting, separator, includeSequenceName, entrySeparator, pageRangeSeparator, includeCategoryHeader });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837703.aspx </remarks>
		/// <param name="shortCitation">string shortCitation</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void NextCitation(string shortCitation)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NextCitation", shortCitation);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198045.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="shortCitation">string shortCitation</param>
		/// <param name="longCitation">optional object longCitation</param>
		/// <param name="longCitationAutoText">optional object longCitationAutoText</param>
		/// <param name="category">optional object category</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Field MarkCitation(NetOffice.WordApi.Range range, string shortCitation, object longCitation, object longCitationAutoText, object category)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkCitation", typeof(NetOffice.WordApi.Field), new object[]{ range, shortCitation, longCitation, longCitationAutoText, category });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198045.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="shortCitation">string shortCitation</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Field MarkCitation(NetOffice.WordApi.Range range, string shortCitation)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkCitation", typeof(NetOffice.WordApi.Field), range, shortCitation);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198045.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="shortCitation">string shortCitation</param>
		/// <param name="longCitation">optional object longCitation</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Field MarkCitation(NetOffice.WordApi.Range range, string shortCitation, object longCitation)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkCitation", typeof(NetOffice.WordApi.Field), range, shortCitation, longCitation);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198045.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="shortCitation">string shortCitation</param>
		/// <param name="longCitation">optional object longCitation</param>
		/// <param name="longCitationAutoText">optional object longCitationAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Field MarkCitation(NetOffice.WordApi.Range range, string shortCitation, object longCitation, object longCitationAutoText)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkCitation", typeof(NetOffice.WordApi.Field), range, shortCitation, longCitation, longCitationAutoText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196874.aspx </remarks>
		/// <param name="shortCitation">string shortCitation</param>
		/// <param name="longCitation">optional object longCitation</param>
		/// <param name="longCitationAutoText">optional object longCitationAutoText</param>
		/// <param name="category">optional object category</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void MarkAllCitations(string shortCitation, object longCitation, object longCitationAutoText, object category)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MarkAllCitations", shortCitation, longCitation, longCitationAutoText, category);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196874.aspx </remarks>
		/// <param name="shortCitation">string shortCitation</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void MarkAllCitations(string shortCitation)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MarkAllCitations", shortCitation);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196874.aspx </remarks>
		/// <param name="shortCitation">string shortCitation</param>
		/// <param name="longCitation">optional object longCitation</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void MarkAllCitations(string shortCitation, object longCitation)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MarkAllCitations", shortCitation, longCitation);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196874.aspx </remarks>
		/// <param name="shortCitation">string shortCitation</param>
		/// <param name="longCitation">optional object longCitation</param>
		/// <param name="longCitationAutoText">optional object longCitationAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void MarkAllCitations(string shortCitation, object longCitation, object longCitationAutoText)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MarkAllCitations", shortCitation, longCitation, longCitationAutoText);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.TableOfAuthorities>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.TableOfAuthorities>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.WordApi.TableOfAuthorities>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.WordApi.TableOfAuthorities>

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.WordApi.TableOfAuthorities> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.WordApi.TableOfAuthorities item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

