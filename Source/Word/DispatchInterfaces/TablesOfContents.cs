using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface TablesOfContents 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.tablesofcontents"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class TablesOfContents : COMObject, IEnumerableProvider<NetOffice.WordApi.TableOfContents>
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
                    _type = typeof(TablesOfContents);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public TablesOfContents(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public TablesOfContents(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TablesOfContents(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TablesOfContents(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TablesOfContents(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TablesOfContents(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TablesOfContents() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TablesOfContents(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.Application"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.Creator"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.Parent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.Count"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.Format"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdTocFormat Format
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdTocFormat>(this, "Format");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Format", value);
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
		public NetOffice.WordApi.TableOfContents this[Int32 index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Item", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, index);
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
		public NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "AddOld", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "AddOld", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "AddOld", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, range, useHeadingStyles);
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
		public NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "AddOld", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, range, useHeadingStyles, upperHeadingLevel);
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
		public NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "AddOld", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel);
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
		public NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "AddOld", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields });
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
		public NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "AddOld", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID });
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
		public NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "AddOld", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers });
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
		public NetOffice.WordApi.TableOfContents AddOld(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "AddOld", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.MarkEntry"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="level">optional object level</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object tableID, object level)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", NetOffice.WordApi.Field.LateBindingApiWrapperType, new object[]{ range, entry, entryAutoText, tableID, level });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.MarkEntry"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", NetOffice.WordApi.Field.LateBindingApiWrapperType, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.MarkEntry"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", NetOffice.WordApi.Field.LateBindingApiWrapperType, range, entry);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.MarkEntry"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", NetOffice.WordApi.Field.LateBindingApiWrapperType, range, entry, entryAutoText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.MarkEntry"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="tableID">optional object tableID</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object tableID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", NetOffice.WordApi.Field.LateBindingApiWrapperType, range, entry, entryAutoText, tableID);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.Add"/> </remarks>
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
		public NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles, object useHyperlinks, object hidePageNumbersInWeb)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles, useHyperlinks, hidePageNumbersInWeb });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.Add"/> </remarks>
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
		public NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles, object useHyperlinks, object hidePageNumbersInWeb, object useOutlineLevels)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles, useHyperlinks, hidePageNumbersInWeb, useOutlineLevels });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.Add"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.Add"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, range, useHeadingStyles);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.Add"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, range, useHeadingStyles, upperHeadingLevel);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.Add"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.Add"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.Add"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.Add"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		/// <param name="upperHeadingLevel">optional object upperHeadingLevel</param>
		/// <param name="lowerHeadingLevel">optional object lowerHeadingLevel</param>
		/// <param name="useFields">optional object useFields</param>
		/// <param name="tableID">optional object tableID</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.Add"/> </remarks>
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
		public NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.Add"/> </remarks>
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
		public NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.TablesOfContents.Add"/> </remarks>
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
		public NetOffice.WordApi.TableOfContents Add(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles, object useHyperlinks)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles, useHyperlinks });
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
		public NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles, object useHyperlinks, object hidePageNumbersInWeb)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles, useHyperlinks, hidePageNumbersInWeb });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, range);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="useHeadingStyles">optional object useHeadingStyles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, range, useHeadingStyles);
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
		public NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, range, useHeadingStyles, upperHeadingLevel);
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
		public NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel);
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
		public NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields });
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
		public NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID });
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
		public NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers });
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
		public NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers });
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
		public NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles });
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
		public NetOffice.WordApi.TableOfContents Add2000(NetOffice.WordApi.Range range, object useHeadingStyles, object upperHeadingLevel, object lowerHeadingLevel, object useFields, object tableID, object rightAlignPageNumbers, object includePageNumbers, object addedStyles, object useHyperlinks)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.TableOfContents>(this, "Add2000", NetOffice.WordApi.TableOfContents.LateBindingApiWrapperType, new object[]{ range, useHeadingStyles, upperHeadingLevel, lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles, useHyperlinks });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.TableOfContents>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.TableOfContents>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
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
        public IEnumerator<NetOffice.WordApi.TableOfContents> GetEnumerator()
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
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}