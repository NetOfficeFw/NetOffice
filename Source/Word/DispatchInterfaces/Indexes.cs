using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface Indexes 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844981.aspx
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Indexes : COMObject ,IEnumerable<NetOffice.WordApi.Index>
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Indexes(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834289.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822602.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195310.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839153.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840322.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdIndexFormat Format
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Format", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdIndexFormat)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Format", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.WordApi.Index this[Int32 index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.WordApi.Index newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Index.LateBindingApiWrapperType) as NetOffice.WordApi.Index;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="headingSeparator">optional object HeadingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object RightAlignPageNumbers</param>
		/// <param name="type">optional object Type</param>
		/// <param name="numberOfColumns">optional object NumberOfColumns</param>
		/// <param name="accentedLetters">optional object AccentedLetters</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns, object accentedLetters)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, headingSeparator, rightAlignPageNumbers, type, numberOfColumns, accentedLetters);
			object returnItem = Invoker.MethodReturn(this, "AddOld", paramsArray);
			NetOffice.WordApi.Index newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Index.LateBindingApiWrapperType) as NetOffice.WordApi.Index;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			object returnItem = Invoker.MethodReturn(this, "AddOld", paramsArray);
			NetOffice.WordApi.Index newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Index.LateBindingApiWrapperType) as NetOffice.WordApi.Index;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="headingSeparator">optional object HeadingSeparator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, headingSeparator);
			object returnItem = Invoker.MethodReturn(this, "AddOld", paramsArray);
			NetOffice.WordApi.Index newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Index.LateBindingApiWrapperType) as NetOffice.WordApi.Index;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="headingSeparator">optional object HeadingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object RightAlignPageNumbers</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, headingSeparator, rightAlignPageNumbers);
			object returnItem = Invoker.MethodReturn(this, "AddOld", paramsArray);
			NetOffice.WordApi.Index newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Index.LateBindingApiWrapperType) as NetOffice.WordApi.Index;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="headingSeparator">optional object HeadingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object RightAlignPageNumbers</param>
		/// <param name="type">optional object Type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, headingSeparator, rightAlignPageNumbers, type);
			object returnItem = Invoker.MethodReturn(this, "AddOld", paramsArray);
			NetOffice.WordApi.Index newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Index.LateBindingApiWrapperType) as NetOffice.WordApi.Index;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="headingSeparator">optional object HeadingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object RightAlignPageNumbers</param>
		/// <param name="type">optional object Type</param>
		/// <param name="numberOfColumns">optional object NumberOfColumns</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index AddOld(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, headingSeparator, rightAlignPageNumbers, type, numberOfColumns);
			object returnItem = Invoker.MethodReturn(this, "AddOld", paramsArray);
			NetOffice.WordApi.Index newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Index.LateBindingApiWrapperType) as NetOffice.WordApi.Index;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="entry">optional object Entry</param>
		/// <param name="entryAutoText">optional object EntryAutoText</param>
		/// <param name="crossReference">optional object CrossReference</param>
		/// <param name="crossReferenceAutoText">optional object CrossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object BookmarkName</param>
		/// <param name="bold">optional object Bold</param>
		/// <param name="italic">optional object Italic</param>
		/// <param name="reading">optional object Reading</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold, object italic, object reading)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName, bold, italic, reading);
			object returnItem = Invoker.MethodReturn(this, "MarkEntry", paramsArray);
			NetOffice.WordApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Field.LateBindingApiWrapperType) as NetOffice.WordApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			object returnItem = Invoker.MethodReturn(this, "MarkEntry", paramsArray);
			NetOffice.WordApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Field.LateBindingApiWrapperType) as NetOffice.WordApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="entry">optional object Entry</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, entry);
			object returnItem = Invoker.MethodReturn(this, "MarkEntry", paramsArray);
			NetOffice.WordApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Field.LateBindingApiWrapperType) as NetOffice.WordApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="entry">optional object Entry</param>
		/// <param name="entryAutoText">optional object EntryAutoText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, entry, entryAutoText);
			object returnItem = Invoker.MethodReturn(this, "MarkEntry", paramsArray);
			NetOffice.WordApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Field.LateBindingApiWrapperType) as NetOffice.WordApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="entry">optional object Entry</param>
		/// <param name="entryAutoText">optional object EntryAutoText</param>
		/// <param name="crossReference">optional object CrossReference</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, entry, entryAutoText, crossReference);
			object returnItem = Invoker.MethodReturn(this, "MarkEntry", paramsArray);
			NetOffice.WordApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Field.LateBindingApiWrapperType) as NetOffice.WordApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="entry">optional object Entry</param>
		/// <param name="entryAutoText">optional object EntryAutoText</param>
		/// <param name="crossReference">optional object CrossReference</param>
		/// <param name="crossReferenceAutoText">optional object CrossReferenceAutoText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, entry, entryAutoText, crossReference, crossReferenceAutoText);
			object returnItem = Invoker.MethodReturn(this, "MarkEntry", paramsArray);
			NetOffice.WordApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Field.LateBindingApiWrapperType) as NetOffice.WordApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="entry">optional object Entry</param>
		/// <param name="entryAutoText">optional object EntryAutoText</param>
		/// <param name="crossReference">optional object CrossReference</param>
		/// <param name="crossReferenceAutoText">optional object CrossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object BookmarkName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName);
			object returnItem = Invoker.MethodReturn(this, "MarkEntry", paramsArray);
			NetOffice.WordApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Field.LateBindingApiWrapperType) as NetOffice.WordApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="entry">optional object Entry</param>
		/// <param name="entryAutoText">optional object EntryAutoText</param>
		/// <param name="crossReference">optional object CrossReference</param>
		/// <param name="crossReferenceAutoText">optional object CrossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object BookmarkName</param>
		/// <param name="bold">optional object Bold</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName, bold);
			object returnItem = Invoker.MethodReturn(this, "MarkEntry", paramsArray);
			NetOffice.WordApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Field.LateBindingApiWrapperType) as NetOffice.WordApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839381.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="entry">optional object Entry</param>
		/// <param name="entryAutoText">optional object EntryAutoText</param>
		/// <param name="crossReference">optional object CrossReference</param>
		/// <param name="crossReferenceAutoText">optional object CrossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object BookmarkName</param>
		/// <param name="bold">optional object Bold</param>
		/// <param name="italic">optional object Italic</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkEntry(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold, object italic)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName, bold, italic);
			object returnItem = Invoker.MethodReturn(this, "MarkEntry", paramsArray);
			NetOffice.WordApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Field.LateBindingApiWrapperType) as NetOffice.WordApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="entry">optional object Entry</param>
		/// <param name="entryAutoText">optional object EntryAutoText</param>
		/// <param name="crossReference">optional object CrossReference</param>
		/// <param name="crossReferenceAutoText">optional object CrossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object BookmarkName</param>
		/// <param name="bold">optional object Bold</param>
		/// <param name="italic">optional object Italic</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold, object italic)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName, bold, italic);
			Invoker.Method(this, "MarkAllEntries", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void MarkAllEntries(NetOffice.WordApi.Range range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			Invoker.Method(this, "MarkAllEntries", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="entry">optional object Entry</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void MarkAllEntries(NetOffice.WordApi.Range range, object entry)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, entry);
			Invoker.Method(this, "MarkAllEntries", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="entry">optional object Entry</param>
		/// <param name="entryAutoText">optional object EntryAutoText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, entry, entryAutoText);
			Invoker.Method(this, "MarkAllEntries", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="entry">optional object Entry</param>
		/// <param name="entryAutoText">optional object EntryAutoText</param>
		/// <param name="crossReference">optional object CrossReference</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, entry, entryAutoText, crossReference);
			Invoker.Method(this, "MarkAllEntries", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="entry">optional object Entry</param>
		/// <param name="entryAutoText">optional object EntryAutoText</param>
		/// <param name="crossReference">optional object CrossReference</param>
		/// <param name="crossReferenceAutoText">optional object CrossReferenceAutoText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, entry, entryAutoText, crossReference, crossReferenceAutoText);
			Invoker.Method(this, "MarkAllEntries", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="entry">optional object Entry</param>
		/// <param name="entryAutoText">optional object EntryAutoText</param>
		/// <param name="crossReference">optional object CrossReference</param>
		/// <param name="crossReferenceAutoText">optional object CrossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object BookmarkName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName);
			Invoker.Method(this, "MarkAllEntries", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837493.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="entry">optional object Entry</param>
		/// <param name="entryAutoText">optional object EntryAutoText</param>
		/// <param name="crossReference">optional object CrossReference</param>
		/// <param name="crossReferenceAutoText">optional object CrossReferenceAutoText</param>
		/// <param name="bookmarkName">optional object BookmarkName</param>
		/// <param name="bold">optional object Bold</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void MarkAllEntries(NetOffice.WordApi.Range range, object entry, object entryAutoText, object crossReference, object crossReferenceAutoText, object bookmarkName, object bold)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, entry, entryAutoText, crossReference, crossReferenceAutoText, bookmarkName, bold);
			Invoker.Method(this, "MarkAllEntries", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841043.aspx
		/// </summary>
		/// <param name="concordanceFileName">string ConcordanceFileName</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void AutoMarkEntries(string concordanceFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(concordanceFileName);
			Invoker.Method(this, "AutoMarkEntries", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="headingSeparator">optional object HeadingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object RightAlignPageNumbers</param>
		/// <param name="type">optional object Type</param>
		/// <param name="numberOfColumns">optional object NumberOfColumns</param>
		/// <param name="accentedLetters">optional object AccentedLetters</param>
		/// <param name="sortBy">optional object SortBy</param>
		/// <param name="indexLanguage">optional object IndexLanguage</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns, object accentedLetters, object sortBy, object indexLanguage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, headingSeparator, rightAlignPageNumbers, type, numberOfColumns, accentedLetters, sortBy, indexLanguage);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Index newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Index.LateBindingApiWrapperType) as NetOffice.WordApi.Index;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Index newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Index.LateBindingApiWrapperType) as NetOffice.WordApi.Index;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="headingSeparator">optional object HeadingSeparator</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, headingSeparator);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Index newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Index.LateBindingApiWrapperType) as NetOffice.WordApi.Index;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="headingSeparator">optional object HeadingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object RightAlignPageNumbers</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, headingSeparator, rightAlignPageNumbers);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Index newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Index.LateBindingApiWrapperType) as NetOffice.WordApi.Index;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="headingSeparator">optional object HeadingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object RightAlignPageNumbers</param>
		/// <param name="type">optional object Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, headingSeparator, rightAlignPageNumbers, type);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Index newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Index.LateBindingApiWrapperType) as NetOffice.WordApi.Index;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="headingSeparator">optional object HeadingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object RightAlignPageNumbers</param>
		/// <param name="type">optional object Type</param>
		/// <param name="numberOfColumns">optional object NumberOfColumns</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, headingSeparator, rightAlignPageNumbers, type, numberOfColumns);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Index newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Index.LateBindingApiWrapperType) as NetOffice.WordApi.Index;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="headingSeparator">optional object HeadingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object RightAlignPageNumbers</param>
		/// <param name="type">optional object Type</param>
		/// <param name="numberOfColumns">optional object NumberOfColumns</param>
		/// <param name="accentedLetters">optional object AccentedLetters</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns, object accentedLetters)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, headingSeparator, rightAlignPageNumbers, type, numberOfColumns, accentedLetters);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Index newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Index.LateBindingApiWrapperType) as NetOffice.WordApi.Index;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193323.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="headingSeparator">optional object HeadingSeparator</param>
		/// <param name="rightAlignPageNumbers">optional object RightAlignPageNumbers</param>
		/// <param name="type">optional object Type</param>
		/// <param name="numberOfColumns">optional object NumberOfColumns</param>
		/// <param name="accentedLetters">optional object AccentedLetters</param>
		/// <param name="sortBy">optional object SortBy</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Index Add(NetOffice.WordApi.Range range, object headingSeparator, object rightAlignPageNumbers, object type, object numberOfColumns, object accentedLetters, object sortBy)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, headingSeparator, rightAlignPageNumbers, type, numberOfColumns, accentedLetters, sortBy);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.Index newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Index.LateBindingApiWrapperType) as NetOffice.WordApi.Index;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.WordApi.Index> Member
        
        /// <summary>
		/// SupportByVersionAttribute Word, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
       public IEnumerator<NetOffice.WordApi.Index> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.WordApi.Index item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Word, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}