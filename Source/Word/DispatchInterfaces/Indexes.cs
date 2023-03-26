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
	/// DispatchInterface Indexes 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.indexes"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class Indexes : COMObject, IEnumerableProvider<NetOffice.WordApi.Index>
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
                    _type = typeof(Indexes);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Indexes(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Indexes(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Indexes(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Indexes(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Indexes(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Indexes(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Indexes() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Indexes(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.Count"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.Format"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdIndexFormat Format
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdIndexFormat>(this, "Format");
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
		public NetOffice.WordApi.Index this[Int32 index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "Item", NetOffice.WordApi.Index.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		/// <param name="numberOfColumns">optional object numberOfColumns</param>
		/// <param name="accentedLetters">optional object accentedLetters</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns, object accentedLetters)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "AddOld", NetOffice.WordApi.Index.LateBindingApiWrapperType, new object[]{ range, headingSeparator, rightAlignPageNumbers, type, numberOfColumns, accentedLetters });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "AddOld", NetOffice.WordApi.Index.LateBindingApiWrapperType, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "AddOld", NetOffice.WordApi.Index.LateBindingApiWrapperType, range, headingSeparator);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "AddOld", NetOffice.WordApi.Index.LateBindingApiWrapperType, range, headingSeparator, rightAlignPageNumbers);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "AddOld", NetOffice.WordApi.Index.LateBindingApiWrapperType, range, headingSeparator, rightAlignPageNumbers, type);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		/// <param name="numberOfColumns">optional object numberOfColumns</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "AddOld", NetOffice.WordApi.Index.LateBindingApiWrapperType, new object[]{ range, headingSeparator, rightAlignPageNumbers, type, numberOfColumns });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.MarkEntry"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		/// <param name="bold">optional object bold</param>
		/// <param name="italic">optional object italic</param>
		/// <param name="reading">optional object reading</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold, object italic, object reading)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", NetOffice.WordApi.Field.LateBindingApiWrapperType, new object[]{ range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName, bold, italic, reading });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.MarkEntry"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.MarkEntry"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.MarkEntry"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.MarkEntry"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", NetOffice.WordApi.Field.LateBindingApiWrapperType, range, entry, entryAutoText, crossReference);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.MarkEntry"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", NetOffice.WordApi.Field.LateBindingApiWrapperType, new object[]{ range, entry, entryAutoText, crossReference, crossReferenceAutoText });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.MarkEntry"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", NetOffice.WordApi.Field.LateBindingApiWrapperType, new object[]{ range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.MarkEntry"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		/// <param name="bold">optional object bold</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", NetOffice.WordApi.Field.LateBindingApiWrapperType, new object[]{ range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName, bold });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.MarkEntry"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		/// <param name="bold">optional object bold</param>
		/// <param name="italic">optional object italic</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold, object italic)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", NetOffice.WordApi.Field.LateBindingApiWrapperType, new object[]{ range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName, bold, italic });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.MarkAllEntries"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		/// <param name="bold">optional object bold</param>
		/// <param name="italic">optional object italic</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold, object italic)
		{
			 Factory.ExecuteMethod(this, "MarkAllEntries", new object[]{ range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName, bold, italic });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.MarkAllEntries"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void MarkAllEntries(NetOffice.WordApi.Range range)
		{
			 Factory.ExecuteMethod(this, "MarkAllEntries", range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.MarkAllEntries"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void MarkAllEntries(NetOffice.WordApi.Range range, object entry)
		{
			 Factory.ExecuteMethod(this, "MarkAllEntries", range, entry);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.MarkAllEntries"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText)
		{
			 Factory.ExecuteMethod(this, "MarkAllEntries", range, entry, entryAutoText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.MarkAllEntries"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference)
		{
			 Factory.ExecuteMethod(this, "MarkAllEntries", range, entry, entryAutoText, crossReference);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.MarkAllEntries"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText)
		{
			 Factory.ExecuteMethod(this, "MarkAllEntries", new object[]{ range, entry, entryAutoText, crossReference, crossReferenceAutoText });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.MarkAllEntries"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName)
		{
			 Factory.ExecuteMethod(this, "MarkAllEntries", new object[]{ range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.MarkAllEntries"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		/// <param name="bold">optional object bold</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold)
		{
			 Factory.ExecuteMethod(this, "MarkAllEntries", new object[]{ range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName, bold });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.AutoMarkEntries"/> </remarks>
		/// <param name="concordanceFileName">string concordanceFileName</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AutoMarkEntries(string concordanceFileName)
		{
			 Factory.ExecuteMethod(this, "AutoMarkEntries", concordanceFileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.Add"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		/// <param name="numberOfColumns">optional object numberOfColumns</param>
		/// <param name="accentedLetters">optional object accentedLetters</param>
		/// <param name="sortBy">optional object sortBy</param>
		/// <param name="indexLanguage">optional object indexLanguage</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns, object accentedLetters, object sortBy, object indexLanguage)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "Add", NetOffice.WordApi.Index.LateBindingApiWrapperType, new object[]{ range, headingSeparator, rightAlignPageNumbers, type, numberOfColumns, accentedLetters, sortBy, indexLanguage });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.Add"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "Add", NetOffice.WordApi.Index.LateBindingApiWrapperType, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.Add"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "Add", NetOffice.WordApi.Index.LateBindingApiWrapperType, range, headingSeparator);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.Add"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "Add", NetOffice.WordApi.Index.LateBindingApiWrapperType, range, headingSeparator, rightAlignPageNumbers);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.Add"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "Add", NetOffice.WordApi.Index.LateBindingApiWrapperType, range, headingSeparator, rightAlignPageNumbers, type);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.Add"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		/// <param name="numberOfColumns">optional object numberOfColumns</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "Add", NetOffice.WordApi.Index.LateBindingApiWrapperType, new object[]{ range, headingSeparator, rightAlignPageNumbers, type, numberOfColumns });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.Add"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		/// <param name="numberOfColumns">optional object numberOfColumns</param>
		/// <param name="accentedLetters">optional object accentedLetters</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns, object accentedLetters)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "Add", NetOffice.WordApi.Index.LateBindingApiWrapperType, new object[]{ range, headingSeparator, rightAlignPageNumbers, type, numberOfColumns, accentedLetters });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Indexes.Add"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		/// <param name="numberOfColumns">optional object numberOfColumns</param>
		/// <param name="accentedLetters">optional object accentedLetters</param>
		/// <param name="sortBy">optional object sortBy</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns, object accentedLetters, object sortBy)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "Add", NetOffice.WordApi.Index.LateBindingApiWrapperType, new object[]{ range, headingSeparator, rightAlignPageNumbers, type, numberOfColumns, accentedLetters, sortBy });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.Index>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.Index>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.WordApi.Index>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.WordApi.Index>

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.WordApi.Index> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.WordApi.Index item in innerEnumerator)
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