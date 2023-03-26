using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface Charts 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class Charts : COMObject, IEnumerableProvider<object>
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
                    _type = typeof(Charts);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Charts(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Charts(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Charts(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Charts(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Charts(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Charts(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Charts() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Charts(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.Application"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.Creator"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.Parent"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.Count"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.HPageBreaks"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.HPageBreaks HPageBreaks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.HPageBreaks>(this, "HPageBreaks", NetOffice.ExcelApi.HPageBreaks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.VPageBreaks"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.VPageBreaks VPageBreaks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.VPageBreaks>(this, "VPageBreaks", NetOffice.ExcelApi.VPageBreaks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.Visible"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Visible
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Visible");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Visible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public object this[object index]
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "_Default", index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Chart Add(object before, object after, object count)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Chart>(this, "Add", NetOffice.ExcelApi.Chart.LateBindingApiWrapperType, before, after, count);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Chart Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Chart>(this, "Add", NetOffice.ExcelApi.Chart.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Chart Add(object before)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Chart>(this, "Add", NetOffice.ExcelApi.Chart.LateBindingApiWrapperType, before);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Chart Add(object before, object after)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Chart>(this, "Add", NetOffice.ExcelApi.Chart.LateBindingApiWrapperType, before, after);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.Copy"/> </remarks>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Copy(object before, object after)
		{
			 Factory.ExecuteMethod(this, "Copy", before, after);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.Copy"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.Copy"/> </remarks>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Copy(object before)
		{
			 Factory.ExecuteMethod(this, "Copy", before);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.Delete"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _Dummy7()
		{
			 Factory.ExecuteMethod(this, "_Dummy7");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.Move"/> </remarks>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Move(object before, object after)
		{
			 Factory.ExecuteMethod(this, "Move", before, after);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.Move"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Move()
		{
			 Factory.ExecuteMethod(this, "Move");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.Move"/> </remarks>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Move(object before)
		{
			 Factory.ExecuteMethod(this, "Move", before);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
		{
			 Factory.ExecuteMethod(this, "_PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="prToFileName">optional object prToFileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
		{
			 Factory.ExecuteMethod(this, "_PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate, prToFileName });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut()
		{
			 Factory.ExecuteMethod(this, "_PrintOut");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from)
		{
			 Factory.ExecuteMethod(this, "_PrintOut", from);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from, object to)
		{
			 Factory.ExecuteMethod(this, "_PrintOut", from, to);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from, object to, object copies)
		{
			 Factory.ExecuteMethod(this, "_PrintOut", from, to, copies);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from, object to, object copies, object preview)
		{
			 Factory.ExecuteMethod(this, "_PrintOut", from, to, copies, preview);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from, object to, object copies, object preview, object activePrinter)
		{
			 Factory.ExecuteMethod(this, "_PrintOut", new object[]{ from, to, copies, preview, activePrinter });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
		{
			 Factory.ExecuteMethod(this, "_PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.PrintPreview"/> </remarks>
		/// <param name="enableChanges">optional object enableChanges</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintPreview(object enableChanges)
		{
			 Factory.ExecuteMethod(this, "PrintPreview", enableChanges);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.PrintPreview"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintPreview()
		{
			 Factory.ExecuteMethod(this, "PrintPreview");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.Select"/> </remarks>
		/// <param name="replace">optional object replace</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Select(object replace)
		{
			 Factory.ExecuteMethod(this, "Select", replace);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.Select"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.PrintOut"/> </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="prToFileName">optional object prToFileName</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate, prToFileName });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.PrintOut"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			 Factory.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.PrintOut"/> </remarks>
		/// <param name="from">optional object from</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from)
		{
			 Factory.ExecuteMethod(this, "PrintOut", from);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.PrintOut"/> </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to)
		{
			 Factory.ExecuteMethod(this, "PrintOut", from, to);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.PrintOut"/> </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object copies)
		{
			 Factory.ExecuteMethod(this, "PrintOut", from, to, copies);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.PrintOut"/> </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object copies, object preview)
		{
			 Factory.ExecuteMethod(this, "PrintOut", from, to, copies, preview);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.PrintOut"/> </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object copies, object preview, object activePrinter)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ from, to, copies, preview, activePrinter });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.PrintOut"/> </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Charts.PrintOut"/> </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate)
		{
			 Factory.ExecuteMethod(this, "__PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void __PrintOut()
		{
			 Factory.ExecuteMethod(this, "__PrintOut");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void __PrintOut(object from)
		{
			 Factory.ExecuteMethod(this, "__PrintOut", from);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void __PrintOut(object from, object to)
		{
			 Factory.ExecuteMethod(this, "__PrintOut", from, to);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void __PrintOut(object from, object to, object copies)
		{
			 Factory.ExecuteMethod(this, "__PrintOut", from, to, copies);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void __PrintOut(object from, object to, object copies, object preview)
		{
			 Factory.ExecuteMethod(this, "__PrintOut", from, to, copies, preview);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void __PrintOut(object from, object to, object copies, object preview, object activePrinter)
		{
			 Factory.ExecuteMethod(this, "__PrintOut", new object[]{ from, to, copies, preview, activePrinter });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public void __PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile)
		{
			 Factory.ExecuteMethod(this, "__PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile });
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.charts.add2"/> </remarks>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		/// <param name="count">optional object count</param>
		/// <param name="newLayout">optional object newLayout</param>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Chart Add2(object before, object after, object count, object newLayout)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Chart>(this, "Add2", NetOffice.ExcelApi.Chart.LateBindingApiWrapperType, before, after, count, newLayout);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.charts.add2"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Chart Add2()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Chart>(this, "Add2", NetOffice.ExcelApi.Chart.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.charts.add2"/> </remarks>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Chart Add2(object before)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Chart>(this, "Add2", NetOffice.ExcelApi.Chart.LateBindingApiWrapperType, before);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.charts.add2"/> </remarks>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Chart Add2(object before, object after)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Chart>(this, "Add2", NetOffice.ExcelApi.Chart.LateBindingApiWrapperType, before, after);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.charts.add2"/> </remarks>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		/// <param name="count">optional object count</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Chart Add2(object before, object after, object count)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Chart>(this, "Add2", NetOffice.ExcelApi.Chart.LateBindingApiWrapperType, before, after, count);
		}

        #endregion

        #region IEnumerableProvider<object>

        ICOMObject IEnumerableProvider<object>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<object>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, true);
        }

        #endregion

        #region IEnumerable<object> Member

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<object> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (object item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}