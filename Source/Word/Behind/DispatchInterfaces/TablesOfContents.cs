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
	/// DispatchInterface TablesOfContents 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838538.aspx </remarks>
	public class TablesOfContents : COMObject, NetOffice.WordApi.TablesOfContents
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
                    _contractType = typeof(NetOffice.WordApi.TablesOfContents);
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
                    _type = typeof(TablesOfContents);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public TablesOfContents() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197427.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197796.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840817.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845238.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839904.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdTocFormat Format
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdTocFormat>(this, "Format");
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
		public virtual NetOffice.WordApi.TableOfContents this[Int32 index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Item", typeof(NetOffice.WordApi.TableOfContents), index);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		/// <param name="addedStyles">optional object addedStyles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "AddOld", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "AddOld", typeof(NetOffice.WordApi.TableOfContents), range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "AddOld", typeof(NetOffice.WordApi.TableOfContents), range, useHeadingStyles);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "AddOld", typeof(NetOffice.WordApi.TableOfContents), range, useHeadingStyles, upperHeadingLevel);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "AddOld", typeof(NetOffice.WordApi.TableOfContents), range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "AddOld", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "AddOld", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "AddOld", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "AddOld", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840232.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="level">optional object level</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object tableID, object level)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", typeof(NetOffice.WordApi.Field), new object[]{ range, entry, entryAutoText, tableID, level });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840232.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", typeof(NetOffice.WordApi.Field), range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840232.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", typeof(NetOffice.WordApi.Field), range, entry);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840232.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", typeof(NetOffice.WordApi.Field), range, entry, entryAutoText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840232.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="tableID">optional object tableID</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object tableID)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", typeof(NetOffice.WordApi.Field), range, entry, entryAutoText, tableID);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		/// <param name="addedStyles">optional object addedStyles</param>
		/// <param name="useHyperlinks">optional object useHyperlinks</param>
		/// <param name="hidePageNumbersInWeb">optional object hidePageNumbersInWeb</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles, object useHyperlinks, object hidePageNumbersInWeb)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles, useHyperlinks, hidePageNumbersInWeb });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		/// <param name="addedStyles">optional object addedStyles</param>
		/// <param name="useHyperlinks">optional object useHyperlinks</param>
		/// <param name="hidePageNumbersInWeb">optional object hidePageNumbersInWeb</param>
		/// <param name="useOutlineLevels">optional object useOutlineLevels</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles, object useHyperlinks, object hidePageNumbersInWeb, object useOutlineLevels)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles, useHyperlinks, hidePageNumbersInWeb, useOutlineLevels });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", typeof(NetOffice.WordApi.TableOfContents), range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", typeof(NetOffice.WordApi.TableOfContents), range, useHeadingStyles);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", typeof(NetOffice.WordApi.TableOfContents), range, useHeadingStyles, upperHeadingLevel);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", typeof(NetOffice.WordApi.TableOfContents), range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		/// <param name="addedStyles">optional object addedStyles</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835785.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		/// <param name="addedStyles">optional object addedStyles</param>
		/// <param name="useHyperlinks">optional object useHyperlinks</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles, object useHyperlinks)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles, useHyperlinks });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		/// <param name="addedStyles">optional object addedStyles</param>
		/// <param name="useHyperlinks">optional object useHyperlinks</param>
		/// <param name="hidePageNumbersInWeb">optional object hidePageNumbersInWeb</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles, object useHyperlinks, object hidePageNumbersInWeb)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles, useHyperlinks, hidePageNumbersInWeb });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", typeof(NetOffice.WordApi.TableOfContents), range);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", typeof(NetOffice.WordApi.TableOfContents), range, useHeadingStyles);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", typeof(NetOffice.WordApi.TableOfContents), range, useHeadingStyles, upperHeadingLevel);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", typeof(NetOffice.WordApi.TableOfContents), range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		/// <param name="addedStyles">optional object addedStyles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		/// <param name="addedStyles">optional object addedStyles</param>
		/// <param name="useHyperlinks">optional object useHyperlinks</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles, object useHyperlinks)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", typeof(NetOffice.WordApi.TableOfContents), new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles, useHyperlinks });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.TableOfContents>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.TableOfContents>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.WordApi.TableOfContents>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.WordApi.TableOfContents>

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.WordApi.TableOfContents> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.WordApi.TableOfContents item in innerEnumerator)
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

