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
	/// DispatchInterface TablesOfFigures 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192739.aspx </remarks>
	public class TablesOfFigures : COMObject , NetOffice.WordApi.TablesOfFigures
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
                    _contractType = typeof(NetOffice.WordApi.TablesOfFigures);
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
                    _type = typeof(TablesOfFigures);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public TablesOfFigures() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193031.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845264.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837538.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838872.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192562.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdTofFormat Format
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdTofFormat>(this, "Format");
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
		public virtual NetOffice.WordApi.TableOfFigures this[Int32 index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "Item", typeof(NetOffice.WordApi.TableOfFigures), index);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
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
		public virtual NetOffice.WordApi.TableOfFigures AddOld(NetOffice.WordApi.Range range, object caption, object includeLabel, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "AddOld", typeof(NetOffice.WordApi.TableOfFigures), new object[]{ range, caption, includeLabel, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures AddOld(NetOffice.WordApi.Range range)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "AddOld", typeof(NetOffice.WordApi.TableOfFigures), range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures AddOld(NetOffice.WordApi.Range range, object caption)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "AddOld", typeof(NetOffice.WordApi.TableOfFigures), range, caption);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures AddOld(NetOffice.WordApi.Range range, object caption, object includeLabel)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "AddOld", typeof(NetOffice.WordApi.TableOfFigures), range, caption, includeLabel);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures AddOld(NetOffice.WordApi.Range range, object caption, object includeLabel, object useHeadingStyles)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "AddOld", typeof(NetOffice.WordApi.TableOfFigures), range, caption, includeLabel, useHeadingStyles);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures AddOld(NetOffice.WordApi.Range range, object caption, object includeLabel, object useHeadingStyles, object upperHeadingLevel)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "AddOld", typeof(NetOffice.WordApi.TableOfFigures), new object[]{ range, caption, includeLabel, useHeadingStyles, upperHeadingLevel });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures AddOld(NetOffice.WordApi.Range range, object caption, object includeLabel, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "AddOld", typeof(NetOffice.WordApi.TableOfFigures), new object[]{ range, caption, includeLabel, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures AddOld(NetOffice.WordApi.Range range, object caption, object includeLabel, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "AddOld", typeof(NetOffice.WordApi.TableOfFigures), new object[]{ range, caption, includeLabel, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures AddOld(NetOffice.WordApi.Range range, object caption, object includeLabel, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "AddOld", typeof(NetOffice.WordApi.TableOfFigures), new object[]{ range, caption, includeLabel, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures AddOld(NetOffice.WordApi.Range range, object caption, object includeLabel, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "AddOld", typeof(NetOffice.WordApi.TableOfFigures), new object[]{ range, caption, includeLabel, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
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
		public virtual NetOffice.WordApi.TableOfFigures AddOld(NetOffice.WordApi.Range range, object caption, object includeLabel, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "AddOld", typeof(NetOffice.WordApi.TableOfFigures), new object[]{ range, caption, includeLabel, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837496.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837496.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837496.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837496.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837496.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835191.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
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
		public virtual NetOffice.WordApi.TableOfFigures Add(NetOffice.WordApi.Range range, object caption, object includeLabel, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles, object useHyperlinks, object hidePageNumbersInWeb)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "Add", typeof(NetOffice.WordApi.TableOfFigures), new object[]{ range, caption, includeLabel, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles, useHyperlinks, hidePageNumbersInWeb });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835191.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures Add(NetOffice.WordApi.Range range)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "Add", typeof(NetOffice.WordApi.TableOfFigures), range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835191.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures Add(NetOffice.WordApi.Range range, object caption)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "Add", typeof(NetOffice.WordApi.TableOfFigures), range, caption);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835191.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures Add(NetOffice.WordApi.Range range, object caption, object includeLabel)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "Add", typeof(NetOffice.WordApi.TableOfFigures), range, caption, includeLabel);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835191.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures Add(NetOffice.WordApi.Range range, object caption, object includeLabel, object useHeadingStyles)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "Add", typeof(NetOffice.WordApi.TableOfFigures), range, caption, includeLabel, useHeadingStyles);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835191.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures Add(NetOffice.WordApi.Range range, object caption, object includeLabel, object useHeadingStyles, object upperHeadingLevel)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "Add", typeof(NetOffice.WordApi.TableOfFigures), new object[]{ range, caption, includeLabel, useHeadingStyles, upperHeadingLevel });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835191.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures Add(NetOffice.WordApi.Range range, object caption, object includeLabel, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "Add", typeof(NetOffice.WordApi.TableOfFigures), new object[]{ range, caption, includeLabel, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835191.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures Add(NetOffice.WordApi.Range range, object caption, object includeLabel, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "Add", typeof(NetOffice.WordApi.TableOfFigures), new object[]{ range, caption, includeLabel, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835191.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures Add(NetOffice.WordApi.Range range, object caption, object includeLabel, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "Add", typeof(NetOffice.WordApi.TableOfFigures), new object[]{ range, caption, includeLabel, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835191.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures Add(NetOffice.WordApi.Range range, object caption, object includeLabel, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "Add", typeof(NetOffice.WordApi.TableOfFigures), new object[]{ range, caption, includeLabel, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835191.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="includePageNumbers">optional object includePageNumbers</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TableOfFigures Add(NetOffice.WordApi.Range range, object caption, object includeLabel, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "Add", typeof(NetOffice.WordApi.TableOfFigures), new object[]{ range, caption, includeLabel, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835191.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
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
		public virtual NetOffice.WordApi.TableOfFigures Add(NetOffice.WordApi.Range range, object caption, object includeLabel, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "Add", typeof(NetOffice.WordApi.TableOfFigures), new object[]{ range, caption, includeLabel, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835191.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="caption">optional object caption</param>
		/// <param name="includeLabel">optional object includeLabel</param>
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
		public virtual NetOffice.WordApi.TableOfFigures Add(NetOffice.WordApi.Range range, object caption, object includeLabel, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles, object useHyperlinks)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfFigures>(this, "Add", typeof(NetOffice.WordApi.TableOfFigures), new object[]{ range, caption, includeLabel, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles, useHyperlinks });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.TableOfFigures>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.TableOfFigures>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.WordApi.TableOfFigures>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.WordApi.TableOfFigures>

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.WordApi.TableOfFigures> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.WordApi.TableOfFigures item in innerEnumerator)
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

