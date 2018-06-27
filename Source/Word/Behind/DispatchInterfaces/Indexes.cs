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
	/// DispatchInterface Indexes 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844981.aspx </remarks>
	public class Indexes : COMObject, NetOffice.WordApi.Indexes
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
                    _contractType = typeof(NetOffice.WordApi.Indexes);
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
                    _type = typeof(Indexes);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Indexes() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834289.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822602.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195310.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839153.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840322.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdIndexFormat Format
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdIndexFormat>(this, "Format");
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
		public virtual NetOffice.WordApi.Index this[Int32 index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "Item", typeof(NetOffice.WordApi.Index), index);
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
		public virtual NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns, object accentedLetters)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "AddOld", typeof(NetOffice.WordApi.Index), new object[]{ range, headingSeparator, rightAlignPageNumbers, type, numberOfColumns, accentedLetters });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "AddOld", typeof(NetOffice.WordApi.Index), range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "AddOld", typeof(NetOffice.WordApi.Index), range, headingSeparator);
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
		public virtual NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "AddOld", typeof(NetOffice.WordApi.Index), range, headingSeparator, rightAlignPageNumbers);
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
		public virtual NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "AddOld", typeof(NetOffice.WordApi.Index), range, headingSeparator, rightAlignPageNumbers, type);
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
		public virtual NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "AddOld", typeof(NetOffice.WordApi.Index), new object[]{ range, headingSeparator, rightAlignPageNumbers, type, numberOfColumns });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx </remarks>
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
		public virtual NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold, object italic, object reading)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", typeof(NetOffice.WordApi.Field), new object[]{ range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName, bold, italic, reading });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", typeof(NetOffice.WordApi.Field), range, entry, entryAutoText, crossReference);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", typeof(NetOffice.WordApi.Field), new object[]{ range, entry, entryAutoText, crossReference, crossReferenceAutoText });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", typeof(NetOffice.WordApi.Field), new object[]{ range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		/// <param name="bold">optional object bold</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", typeof(NetOffice.WordApi.Field), new object[]{ range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName, bold });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx </remarks>
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
		public virtual NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold, object italic)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "MarkEntry", typeof(NetOffice.WordApi.Field), new object[]{ range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName, bold, italic });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		/// <param name="bold">optional object bold</param>
		/// <param name="italic">optional object italic</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold, object italic)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MarkAllEntries", new object[]{ range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName, bold, italic });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void MarkAllEntries(NetOffice.WordApi.Range range)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MarkAllEntries", range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void MarkAllEntries(NetOffice.WordApi.Range range, object entry)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MarkAllEntries", range, entry);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MarkAllEntries", range, entry, entryAutoText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MarkAllEntries", range, entry, entryAutoText, crossReference);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MarkAllEntries", new object[]{ range, entry, entryAutoText, crossReference, crossReferenceAutoText });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MarkAllEntries", new object[]{ range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="entry">optional object entry</param>
		/// <param name="entryAutoText">optional object entryAutoText</param>
		/// <param name="crossReference">optional object crossReference</param>
		/// <param name="crossReferenceAutoText">optional object crossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object bookmarkName</param>
		/// <param name="bold">optional object bold</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MarkAllEntries", new object[]{ range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName, bold });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841043.aspx </remarks>
		/// <param name="concordanceFileName">string concordanceFileName</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void AutoMarkEntries(string concordanceFileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AutoMarkEntries", concordanceFileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		/// <param name="numberOfColumns">optional object numberOfColumns</param>
		/// <param name="accentedLetters">optional object accentedLetters</param>
		/// <param name="sortBy">optional object sortBy</param>
		/// <param name="indexLanguage">optional object indexLanguage</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns, object accentedLetters, object sortBy, object indexLanguage)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "Add", typeof(NetOffice.WordApi.Index), new object[]{ range, headingSeparator, rightAlignPageNumbers, type, numberOfColumns, accentedLetters, sortBy, indexLanguage });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "Add", typeof(NetOffice.WordApi.Index), range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "Add", typeof(NetOffice.WordApi.Index), range, headingSeparator);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "Add", typeof(NetOffice.WordApi.Index), range, headingSeparator, rightAlignPageNumbers);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "Add", typeof(NetOffice.WordApi.Index), range, headingSeparator, rightAlignPageNumbers, type);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		/// <param name="numberOfColumns">optional object numberOfColumns</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "Add", typeof(NetOffice.WordApi.Index), new object[]{ range, headingSeparator, rightAlignPageNumbers, type, numberOfColumns });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		/// <param name="numberOfColumns">optional object numberOfColumns</param>
		/// <param name="accentedLetters">optional object accentedLetters</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns, object accentedLetters)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "Add", typeof(NetOffice.WordApi.Index), new object[]{ range, headingSeparator, rightAlignPageNumbers, type, numberOfColumns, accentedLetters });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		/// <param name="headingSeparator">optional object headingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object rightAlignPageNumbers</param>
		/// <param name="type">optional object type</param>
		/// <param name="numberOfColumns">optional object numberOfColumns</param>
		/// <param name="accentedLetters">optional object accentedLetters</param>
		/// <param name="sortBy">optional object sortBy</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns, object accentedLetters, object sortBy)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Index>(this, "Add", typeof(NetOffice.WordApi.Index), new object[]{ range, headingSeparator, rightAlignPageNumbers, type, numberOfColumns, accentedLetters, sortBy });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.Index>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.Index>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
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
        public virtual IEnumerator<NetOffice.WordApi.Index> GetEnumerator()
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
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

