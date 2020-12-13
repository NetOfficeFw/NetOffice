using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// TextRange2
	/// </summary>
	[SyntaxBypass]
 	public class TextRange2_ : _IMsoDispObj
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public TextRange2_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public TextRange2_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange2_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange2_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange2_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange2_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange2_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange2_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="start">optional Int32 start</param>
		/// <param name="length">optional Int32 length</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Paragraphs"/>
		[SupportByVersion("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.TextRange2 get_Paragraphs(object start, object length)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Paragraphs", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, start, length);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Alias for get_Paragraphs
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Paragraphs"/> </remarks>
		/// <param name="start">optional Int32 start</param>
		/// <param name="length">optional Int32 length</param>
		[SupportByVersion("Office", 12,14,15,16), Redirect("get_Paragraphs")]
		public NetOffice.OfficeApi.TextRange2 Paragraphs(object start, object length)
		{
			return get_Paragraphs(start, length);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="start">optional Int32 start</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Paragraphs"/>
		[SupportByVersion("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.TextRange2 get_Paragraphs(object start)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Paragraphs", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Alias for get_Paragraphs
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Paragraphs"/> </remarks>
		/// <param name="start">optional Int32 start</param>
		[SupportByVersion("Office", 12,14,15,16), Redirect("get_Paragraphs")]
		public NetOffice.OfficeApi.TextRange2 Paragraphs(object start)
		{
			return get_Paragraphs(start);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="start">optional Int32 start</param>
		/// <param name="length">optional Int32 length</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Sentences"/>
		[SupportByVersion("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.TextRange2 get_Sentences(object start, object length)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Sentences", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, start, length);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Alias for get_Sentences
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Sentences"/> </remarks>
		/// <param name="start">optional Int32 start</param>
		/// <param name="length">optional Int32 length</param>
		[SupportByVersion("Office", 12,14,15,16), Redirect("get_Sentences")]
		public NetOffice.OfficeApi.TextRange2 Sentences(object start, object length)
		{
			return get_Sentences(start, length);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="start">optional Int32 start</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Sentences"/>
		[SupportByVersion("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.TextRange2 get_Sentences(object start)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Sentences", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Alias for get_Sentences
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Sentences"/> </remarks>
		/// <param name="start">optional Int32 start</param>
		[SupportByVersion("Office", 12,14,15,16), Redirect("get_Sentences")]
		public NetOffice.OfficeApi.TextRange2 Sentences(object start)
		{
			return get_Sentences(start);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="start">optional Int32 start</param>
		/// <param name="length">optional Int32 length</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Words"/>
		[SupportByVersion("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.TextRange2 get_Words(object start, object length)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Words", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, start, length);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Alias for get_Words
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Words"/> </remarks>
		/// <param name="start">optional Int32 start</param>
		/// <param name="length">optional Int32 length</param>
		[SupportByVersion("Office", 12,14,15,16), Redirect("get_Words")]
		public NetOffice.OfficeApi.TextRange2 Words(object start, object length)
		{
			return get_Words(start, length);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="start">optional Int32 start</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Words"/>
		[SupportByVersion("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.TextRange2 get_Words(object start)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Words", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Alias for get_Words
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Words"/> </remarks>
		/// <param name="start">optional Int32 start</param>
		[SupportByVersion("Office", 12,14,15,16), Redirect("get_Words")]
		public NetOffice.OfficeApi.TextRange2 Words(object start)
		{
			return get_Words(start);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="start">optional Int32 start</param>
		/// <param name="length">optional Int32 length</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Characters"/>
		[SupportByVersion("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.TextRange2 get_Characters(object start, object length)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Characters", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, start, length);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Alias for get_Characters
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Characters"/> </remarks>
		/// <param name="start">optional Int32 start</param>
		/// <param name="length">optional Int32 length</param>
		[SupportByVersion("Office", 12,14,15,16), Redirect("get_Characters")]
		public NetOffice.OfficeApi.TextRange2 Characters(object start, object length)
		{
			return get_Characters(start, length);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="start">optional Int32 start</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Characters"/>
		[SupportByVersion("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.TextRange2 get_Characters(object start)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Characters", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Alias for get_Characters
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Characters"/> </remarks>
		/// <param name="start">optional Int32 start</param>
		[SupportByVersion("Office", 12,14,15,16), Redirect("get_Characters")]
		public NetOffice.OfficeApi.TextRange2 Characters(object start)
		{
			return get_Characters(start);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="start">optional Int32 start</param>
		/// <param name="length">optional Int32 length</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Lines"/>
		[SupportByVersion("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.TextRange2 get_Lines(object start, object length)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Lines", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, start, length);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Alias for get_Lines
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Lines"/> </remarks>
		/// <param name="start">optional Int32 start</param>
		/// <param name="length">optional Int32 length</param>
		[SupportByVersion("Office", 12,14,15,16), Redirect("get_Lines")]
		public NetOffice.OfficeApi.TextRange2 Lines(object start, object length)
		{
			return get_Lines(start, length);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="start">optional Int32 start</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Lines"/>
		[SupportByVersion("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.TextRange2 get_Lines(object start)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Lines", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Alias for get_Lines
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Lines"/> </remarks>
		/// <param name="start">optional Int32 start</param>
		[SupportByVersion("Office", 12,14,15,16), Redirect("get_Lines")]
		public NetOffice.OfficeApi.TextRange2 Lines(object start)
		{
			return get_Lines(start);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="start">optional Int32 start</param>
		/// <param name="length">optional Int32 length</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Runs"/>
		[SupportByVersion("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.TextRange2 get_Runs(object start, object length)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Runs", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, start, length);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Alias for get_Runs
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Runs"/> </remarks>
		/// <param name="start">optional Int32 start</param>
		/// <param name="length">optional Int32 length</param>
		[SupportByVersion("Office", 12,14,15,16), Redirect("get_Runs")]
		public NetOffice.OfficeApi.TextRange2 Runs(object start, object length)
		{
			return get_Runs(start, length);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="start">optional Int32 start</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Runs"/>
		[SupportByVersion("Office", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.TextRange2 get_Runs(object start)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Runs", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Alias for get_Runs
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Runs"/> </remarks>
		/// <param name="start">optional Int32 start</param>
		[SupportByVersion("Office", 12,14,15,16), Redirect("get_Runs")]
		public NetOffice.OfficeApi.TextRange2 Runs(object start)
		{
			return get_Runs(start);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="start">optional Int32 start</param>
		/// <param name="length">optional Int32 length</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.MathZones"/>
		[SupportByVersion("Office", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.TextRange2 get_MathZones(object start, object length)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "MathZones", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, start, length);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Alias for get_MathZones
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.MathZones"/> </remarks>
		/// <param name="start">optional Int32 start</param>
		/// <param name="length">optional Int32 length</param>
		[SupportByVersion("Office", 14,15,16), Redirect("get_MathZones")]
		public NetOffice.OfficeApi.TextRange2 MathZones(object start, object length)
		{
			return get_MathZones(start, length);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="start">optional Int32 start</param>
		/// MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.MathZones"/>
		[SupportByVersion("Office", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.TextRange2 get_MathZones(object start)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "MathZones", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, start);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Alias for get_MathZones
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.MathZones"/> </remarks>
		/// <param name="start">optional Int32 start</param>
		[SupportByVersion("Office", 14,15,16), Redirect("get_MathZones")]
		public NetOffice.OfficeApi.TextRange2 MathZones(object start)
		{
			return get_MathZones(start);
		}

		#endregion

		#region Methods

		#endregion
	}

	/// <summary>
	/// DispatchInterface TextRange2 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2"/> </remarks>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class TextRange2 : TextRange2_, IEnumerableProvider<NetOffice.OfficeApi.TextRange2>
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
                    _type = typeof(TextRange2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public TextRange2(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public TextRange2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange2() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextRange2(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Text"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Count"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Parent"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Paragraphs"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 Paragraphs
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Paragraphs", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Sentences"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 Sentences
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Sentences", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Words"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 Words
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Words", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Characters"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 Characters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Characters", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Lines"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 Lines
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Lines", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Runs"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 Runs
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Runs", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.ParagraphFormat"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.ParagraphFormat2 ParagraphFormat
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ParagraphFormat2>(this, "ParagraphFormat", NetOffice.OfficeApi.ParagraphFormat2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Font"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Font2 Font
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Font2>(this, "Font", NetOffice.OfficeApi.Font2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Length"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 Length
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Length");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Start"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public Int32 Start
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Start");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.BoundLeft"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public Single BoundLeft
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BoundLeft");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.BoundTop"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public Single BoundTop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BoundTop");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.BoundWidth"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public Single BoundWidth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BoundWidth");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.BoundHeight"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public Single BoundHeight
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BoundHeight");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.LanguageID"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoLanguageID LanguageID
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoLanguageID>(this, "LanguageID");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LanguageID", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.MathZones"/> </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.TextRange2 MathZones
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "MathZones", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Office", 12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OfficeApi.TextRange2 this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Item", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.TrimText"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 TrimText()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "TrimText", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.InsertAfter"/> </remarks>
		/// <param name="newText">optional string NewText = </param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 InsertAfter(object newText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "InsertAfter", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, newText);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.InsertAfter"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 InsertAfter()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "InsertAfter", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.InsertBefore"/> </remarks>
		/// <param name="newText">optional string NewText = </param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 InsertBefore(object newText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "InsertBefore", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, newText);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.InsertBefore"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 InsertBefore()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "InsertBefore", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.InsertSymbol"/> </remarks>
		/// <param name="fontName">string fontName</param>
		/// <param name="charNumber">Int32 charNumber</param>
		/// <param name="unicode">optional NetOffice.OfficeApi.Enums.MsoTriState Unicode = 0</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 InsertSymbol(string fontName, Int32 charNumber, object unicode)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "InsertSymbol", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, fontName, charNumber, unicode);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.InsertSymbol"/> </remarks>
		/// <param name="fontName">string fontName</param>
		/// <param name="charNumber">Int32 charNumber</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 InsertSymbol(string fontName, Int32 charNumber)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "InsertSymbol", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, fontName, charNumber);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Select"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Cut"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public void Cut()
		{
			 Factory.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Copy"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Delete"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Paste"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 Paste()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Paste", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.PasteSpecial"/> </remarks>
		/// <param name="format">NetOffice.OfficeApi.Enums.MsoClipboardFormat format</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 PasteSpecial(NetOffice.OfficeApi.Enums.MsoClipboardFormat format)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "PasteSpecial", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, format);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.ChangeCase"/> </remarks>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoTextChangeCase type</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void ChangeCase(NetOffice.OfficeApi.Enums.MsoTextChangeCase type)
		{
			 Factory.ExecuteMethod(this, "ChangeCase", type);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.AddPeriods"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public void AddPeriods()
		{
			 Factory.ExecuteMethod(this, "AddPeriods");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.RemovePeriods"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public void RemovePeriods()
		{
			 Factory.ExecuteMethod(this, "RemovePeriods");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Find"/> </remarks>
		/// <param name="findWhat">string findWhat</param>
		/// <param name="after">optional Int32 After = 0</param>
		/// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
		/// <param name="wholeWords">optional NetOffice.OfficeApi.Enums.MsoTriState WholeWords = 0</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 Find(string findWhat, object after, object matchCase, object wholeWords)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Find", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, findWhat, after, matchCase, wholeWords);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Find"/> </remarks>
		/// <param name="findWhat">string findWhat</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 Find(string findWhat)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Find", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, findWhat);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Find"/> </remarks>
		/// <param name="findWhat">string findWhat</param>
		/// <param name="after">optional Int32 After = 0</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 Find(string findWhat, object after)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Find", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, findWhat, after);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Find"/> </remarks>
		/// <param name="findWhat">string findWhat</param>
		/// <param name="after">optional Int32 After = 0</param>
		/// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 Find(string findWhat, object after, object matchCase)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Find", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, findWhat, after, matchCase);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Replace"/> </remarks>
		/// <param name="findWhat">string findWhat</param>
		/// <param name="replaceWhat">string replaceWhat</param>
		/// <param name="after">optional Int32 After = 0</param>
		/// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
		/// <param name="wholeWords">optional NetOffice.OfficeApi.Enums.MsoTriState WholeWords = 0</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 Replace(string findWhat, string replaceWhat, object after, object matchCase, object wholeWords)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Replace", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, new object[]{ findWhat, replaceWhat, after, matchCase, wholeWords });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Replace"/> </remarks>
		/// <param name="findWhat">string findWhat</param>
		/// <param name="replaceWhat">string replaceWhat</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 Replace(string findWhat, string replaceWhat)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Replace", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, findWhat, replaceWhat);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Replace"/> </remarks>
		/// <param name="findWhat">string findWhat</param>
		/// <param name="replaceWhat">string replaceWhat</param>
		/// <param name="after">optional Int32 After = 0</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 Replace(string findWhat, string replaceWhat, object after)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Replace", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, findWhat, replaceWhat, after);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.Replace"/> </remarks>
		/// <param name="findWhat">string findWhat</param>
		/// <param name="replaceWhat">string replaceWhat</param>
		/// <param name="after">optional Int32 After = 0</param>
		/// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 Replace(string findWhat, string replaceWhat, object after, object matchCase)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Replace", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, findWhat, replaceWhat, after, matchCase);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.RotatedBounds"/> </remarks>
		/// <param name="x1">Single x1</param>
		/// <param name="y1">Single y1</param>
		/// <param name="x2">Single x2</param>
		/// <param name="y2">Single y2</param>
		/// <param name="x3">Single x3</param>
		/// <param name="y3">Single y3</param>
		/// <param name="x4">Single x4</param>
		/// <param name="y4">Single y4</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void RotatedBounds(out Single x1, out Single y1, out Single x2, out Single y2, out Single x3, out Single y3, out Single x4, out Single y4)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,true,true,true,true,true,true);
			x1 = 0;
			y1 = 0;
			x2 = 0;
			y2 = 0;
			x3 = 0;
			y3 = 0;
			x4 = 0;
			y4 = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(x1, y1, x2, y2, x3, y3, x4, y4);
			Invoker.Method(this, "RotatedBounds", paramsArray, modifiers);
			x1 = (Single)paramsArray[0];
			y1 = (Single)paramsArray[1];
			x2 = (Single)paramsArray[2];
			y2 = (Single)paramsArray[3];
			x3 = (Single)paramsArray[4];
			y3 = (Single)paramsArray[5];
			x4 = (Single)paramsArray[6];
			y4 = (Single)paramsArray[7];
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.RtlRun"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public void RtlRun()
		{
			 Factory.ExecuteMethod(this, "RtlRun");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.LtrRun"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public void LtrRun()
		{
			 Factory.ExecuteMethod(this, "LtrRun");
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.InsertChartField"/> </remarks>
		/// <param name="chartFieldType">NetOffice.OfficeApi.Enums.MsoChartFieldType chartFieldType</param>
		/// <param name="formula">optional string Formula = </param>
		/// <param name="position">optional Int32 Position = -1</param>
		[SupportByVersion("Office", 15, 16)]
		public NetOffice.OfficeApi.TextRange2 InsertChartField(NetOffice.OfficeApi.Enums.MsoChartFieldType chartFieldType, object formula, object position)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "InsertChartField", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, chartFieldType, formula, position);
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.InsertChartField"/> </remarks>
		/// <param name="chartFieldType">NetOffice.OfficeApi.Enums.MsoChartFieldType chartFieldType</param>
		[CustomMethod]
		[SupportByVersion("Office", 15, 16)]
		public NetOffice.OfficeApi.TextRange2 InsertChartField(NetOffice.OfficeApi.Enums.MsoChartFieldType chartFieldType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "InsertChartField", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, chartFieldType);
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextRange2.InsertChartField"/> </remarks>
		/// <param name="chartFieldType">NetOffice.OfficeApi.Enums.MsoChartFieldType chartFieldType</param>
		/// <param name="formula">optional string Formula = </param>
		[CustomMethod]
		[SupportByVersion("Office", 15, 16)]
		public NetOffice.OfficeApi.TextRange2 InsertChartField(NetOffice.OfficeApi.Enums.MsoChartFieldType chartFieldType, object formula)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "InsertChartField", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType, chartFieldType, formula);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.TextRange2>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.TextRange2>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.TextRange2>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.TextRange2>

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.TextRange2> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.TextRange2 item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}