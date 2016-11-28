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
	/// DispatchInterface TablesOfAuthorities 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837712.aspx
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class TablesOfAuthorities : COMObject ,IEnumerable<NetOffice.WordApi.TableOfAuthorities>
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
                    _type = typeof(TablesOfAuthorities);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public TablesOfAuthorities(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TablesOfAuthorities(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TablesOfAuthorities(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TablesOfAuthorities(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TablesOfAuthorities(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TablesOfAuthorities() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TablesOfAuthorities(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820743.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845059.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838690.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837691.aspx
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
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839360.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdToaFormat Format
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Format", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdToaFormat)intReturnItem;
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
		public NetOffice.WordApi.TableOfAuthorities this[Int32 index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.WordApi.TableOfAuthorities newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.TableOfAuthorities.LateBindingApiWrapperType) as NetOffice.WordApi.TableOfAuthorities;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="category">optional object Category</param>
		/// <param name="bookmark">optional object Bookmark</param>
		/// <param name="passim">optional object Passim</param>
		/// <param name="keepEntryFormatting">optional object KeepEntryFormatting</param>
		/// <param name="separator">optional object Separator</param>
		/// <param name="includeSequenceName">optional object IncludeSequenceName</param>
		/// <param name="entrySeparator">optional object EntrySeparator</param>
		/// <param name="pageRangeSeparator">optional object PageRangeSeparator</param>
		/// <param name="includeCategoryHeader">optional object IncludeCategoryHeader</param>
		/// <param name="pageNumberSeparator">optional object PageNumberSeparator</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting, object separator, object includeSequenceName, object entrySeparator, object pageRangeSeparator, object includeCategoryHeader, object pageNumberSeparator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, category, bookmark, passim, keepEntryFormatting, separator, includeSequenceName, entrySeparator, pageRangeSeparator, includeCategoryHeader, pageNumberSeparator);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.TableOfAuthorities newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.TableOfAuthorities.LateBindingApiWrapperType) as NetOffice.WordApi.TableOfAuthorities;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.TableOfAuthorities newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.TableOfAuthorities.LateBindingApiWrapperType) as NetOffice.WordApi.TableOfAuthorities;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="category">optional object Category</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, category);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.TableOfAuthorities newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.TableOfAuthorities.LateBindingApiWrapperType) as NetOffice.WordApi.TableOfAuthorities;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="category">optional object Category</param>
		/// <param name="bookmark">optional object Bookmark</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, category, bookmark);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.TableOfAuthorities newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.TableOfAuthorities.LateBindingApiWrapperType) as NetOffice.WordApi.TableOfAuthorities;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="category">optional object Category</param>
		/// <param name="bookmark">optional object Bookmark</param>
		/// <param name="passim">optional object Passim</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, category, bookmark, passim);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.TableOfAuthorities newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.TableOfAuthorities.LateBindingApiWrapperType) as NetOffice.WordApi.TableOfAuthorities;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="category">optional object Category</param>
		/// <param name="bookmark">optional object Bookmark</param>
		/// <param name="passim">optional object Passim</param>
		/// <param name="keepEntryFormatting">optional object KeepEntryFormatting</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, category, bookmark, passim, keepEntryFormatting);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.TableOfAuthorities newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.TableOfAuthorities.LateBindingApiWrapperType) as NetOffice.WordApi.TableOfAuthorities;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="category">optional object Category</param>
		/// <param name="bookmark">optional object Bookmark</param>
		/// <param name="passim">optional object Passim</param>
		/// <param name="keepEntryFormatting">optional object KeepEntryFormatting</param>
		/// <param name="separator">optional object Separator</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting, object separator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, category, bookmark, passim, keepEntryFormatting, separator);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.TableOfAuthorities newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.TableOfAuthorities.LateBindingApiWrapperType) as NetOffice.WordApi.TableOfAuthorities;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="category">optional object Category</param>
		/// <param name="bookmark">optional object Bookmark</param>
		/// <param name="passim">optional object Passim</param>
		/// <param name="keepEntryFormatting">optional object KeepEntryFormatting</param>
		/// <param name="separator">optional object Separator</param>
		/// <param name="includeSequenceName">optional object IncludeSequenceName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting, object separator, object includeSequenceName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, category, bookmark, passim, keepEntryFormatting, separator, includeSequenceName);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.TableOfAuthorities newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.TableOfAuthorities.LateBindingApiWrapperType) as NetOffice.WordApi.TableOfAuthorities;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="category">optional object Category</param>
		/// <param name="bookmark">optional object Bookmark</param>
		/// <param name="passim">optional object Passim</param>
		/// <param name="keepEntryFormatting">optional object KeepEntryFormatting</param>
		/// <param name="separator">optional object Separator</param>
		/// <param name="includeSequenceName">optional object IncludeSequenceName</param>
		/// <param name="entrySeparator">optional object EntrySeparator</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting, object separator, object includeSequenceName, object entrySeparator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, category, bookmark, passim, keepEntryFormatting, separator, includeSequenceName, entrySeparator);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.TableOfAuthorities newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.TableOfAuthorities.LateBindingApiWrapperType) as NetOffice.WordApi.TableOfAuthorities;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="category">optional object Category</param>
		/// <param name="bookmark">optional object Bookmark</param>
		/// <param name="passim">optional object Passim</param>
		/// <param name="keepEntryFormatting">optional object KeepEntryFormatting</param>
		/// <param name="separator">optional object Separator</param>
		/// <param name="includeSequenceName">optional object IncludeSequenceName</param>
		/// <param name="entrySeparator">optional object EntrySeparator</param>
		/// <param name="pageRangeSeparator">optional object PageRangeSeparator</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting, object separator, object includeSequenceName, object entrySeparator, object pageRangeSeparator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, category, bookmark, passim, keepEntryFormatting, separator, includeSequenceName, entrySeparator, pageRangeSeparator);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.TableOfAuthorities newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.TableOfAuthorities.LateBindingApiWrapperType) as NetOffice.WordApi.TableOfAuthorities;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822964.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="category">optional object Category</param>
		/// <param name="bookmark">optional object Bookmark</param>
		/// <param name="passim">optional object Passim</param>
		/// <param name="keepEntryFormatting">optional object KeepEntryFormatting</param>
		/// <param name="separator">optional object Separator</param>
		/// <param name="includeSequenceName">optional object IncludeSequenceName</param>
		/// <param name="entrySeparator">optional object EntrySeparator</param>
		/// <param name="pageRangeSeparator">optional object PageRangeSeparator</param>
		/// <param name="includeCategoryHeader">optional object IncludeCategoryHeader</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TableOfAuthorities Add(NetOffice.WordApi.Range range, object category, object bookmark, object passim, object keepEntryFormatting, object separator, object includeSequenceName, object entrySeparator, object pageRangeSeparator, object includeCategoryHeader)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, category, bookmark, passim, keepEntryFormatting, separator, includeSequenceName, entrySeparator, pageRangeSeparator, includeCategoryHeader);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.TableOfAuthorities newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.TableOfAuthorities.LateBindingApiWrapperType) as NetOffice.WordApi.TableOfAuthorities;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837703.aspx
		/// </summary>
		/// <param name="shortCitation">string ShortCitation</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void NextCitation(string shortCitation)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shortCitation);
			Invoker.Method(this, "NextCitation", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198045.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="shortCitation">string ShortCitation</param>
		/// <param name="longCitation">optional object LongCitation</param>
		/// <param name="longCitationAutoText">optional object LongCitationAutoText</param>
		/// <param name="category">optional object Category</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkCitation(NetOffice.WordApi.Range range, string shortCitation, object longCitation, object longCitationAutoText, object category)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, shortCitation, longCitation, longCitationAutoText, category);
			object returnItem = Invoker.MethodReturn(this, "MarkCitation", paramsArray);
			NetOffice.WordApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Field.LateBindingApiWrapperType) as NetOffice.WordApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198045.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="shortCitation">string ShortCitation</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkCitation(NetOffice.WordApi.Range range, string shortCitation)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, shortCitation);
			object returnItem = Invoker.MethodReturn(this, "MarkCitation", paramsArray);
			NetOffice.WordApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Field.LateBindingApiWrapperType) as NetOffice.WordApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198045.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="shortCitation">string ShortCitation</param>
		/// <param name="longCitation">optional object LongCitation</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkCitation(NetOffice.WordApi.Range range, string shortCitation, object longCitation)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, shortCitation, longCitation);
			object returnItem = Invoker.MethodReturn(this, "MarkCitation", paramsArray);
			NetOffice.WordApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Field.LateBindingApiWrapperType) as NetOffice.WordApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198045.aspx
		/// </summary>
		/// <param name="range">NetOffice.WordApi.Range Range</param>
		/// <param name="shortCitation">string ShortCitation</param>
		/// <param name="longCitation">optional object LongCitation</param>
		/// <param name="longCitationAutoText">optional object LongCitationAutoText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field MarkCitation(NetOffice.WordApi.Range range, string shortCitation, object longCitation, object longCitationAutoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(range, shortCitation, longCitation, longCitationAutoText);
			object returnItem = Invoker.MethodReturn(this, "MarkCitation", paramsArray);
			NetOffice.WordApi.Field newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Field.LateBindingApiWrapperType) as NetOffice.WordApi.Field;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196874.aspx
		/// </summary>
		/// <param name="shortCitation">string ShortCitation</param>
		/// <param name="longCitation">optional object LongCitation</param>
		/// <param name="longCitationAutoText">optional object LongCitationAutoText</param>
		/// <param name="category">optional object Category</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void MarkAllCitations(string shortCitation, object longCitation, object longCitationAutoText, object category)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shortCitation, longCitation, longCitationAutoText, category);
			Invoker.Method(this, "MarkAllCitations", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196874.aspx
		/// </summary>
		/// <param name="shortCitation">string ShortCitation</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void MarkAllCitations(string shortCitation)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shortCitation);
			Invoker.Method(this, "MarkAllCitations", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196874.aspx
		/// </summary>
		/// <param name="shortCitation">string ShortCitation</param>
		/// <param name="longCitation">optional object LongCitation</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void MarkAllCitations(string shortCitation, object longCitation)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shortCitation, longCitation);
			Invoker.Method(this, "MarkAllCitations", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196874.aspx
		/// </summary>
		/// <param name="shortCitation">string ShortCitation</param>
		/// <param name="longCitation">optional object LongCitation</param>
		/// <param name="longCitationAutoText">optional object LongCitationAutoText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void MarkAllCitations(string shortCitation, object longCitation, object longCitationAutoText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(shortCitation, longCitation, longCitationAutoText);
			Invoker.Method(this, "MarkAllCitations", paramsArray);
		}

		#endregion

       #region IEnumerable<NetOffice.WordApi.TableOfAuthorities> Member
        
        /// <summary>
		/// SupportByVersionAttribute Word, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
       public IEnumerator<NetOffice.WordApi.TableOfAuthorities> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.WordApi.TableOfAuthorities item in innerEnumerator)
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