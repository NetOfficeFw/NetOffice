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
	/// DispatchInterface Worksheets 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class Worksheets : COMObject, IEnumerableProvider<object>
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
                    _type = typeof(Worksheets);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Worksheets(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Worksheets(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Worksheets(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Worksheets(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Worksheets(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Worksheets(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Worksheets() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Worksheets(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Count"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.HPageBreaks"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.VPageBreaks"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Visible"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Add"/> </remarks>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		/// <param name="count">optional object count</param>
		/// <param name="type">optional object type</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Add(object before, object after, object count, object type)
		{
			return Factory.ExecuteVariantMethodGet(this, "Add", before, after, count, type);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Add"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Add()
		{
			return Factory.ExecuteVariantMethodGet(this, "Add");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Add"/> </remarks>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Add(object before)
		{
			return Factory.ExecuteVariantMethodGet(this, "Add", before);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Add"/> </remarks>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Add(object before, object after)
		{
			return Factory.ExecuteVariantMethodGet(this, "Add", before, after);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Add"/> </remarks>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		/// <param name="count">optional object count</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Add(object before, object after, object count)
		{
			return Factory.ExecuteVariantMethodGet(this, "Add", before, after, count);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Copy"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Copy"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Copy"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Delete"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.FillAcrossSheets"/> </remarks>
		/// <param name="range">NetOffice.ExcelApi.Range range</param>
		/// <param name="type">optional NetOffice.ExcelApi.Enums.XlFillWith Type = -4104</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void FillAcrossSheets(NetOffice.ExcelApi.Range range, object type)
		{
			 Factory.ExecuteMethod(this, "FillAcrossSheets", range, type);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.FillAcrossSheets"/> </remarks>
		/// <param name="range">NetOffice.ExcelApi.Range range</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void FillAcrossSheets(NetOffice.ExcelApi.Range range)
		{
			 Factory.ExecuteMethod(this, "FillAcrossSheets", range);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Move"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Move"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Move()
		{
			 Factory.ExecuteMethod(this, "Move");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Move"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.PrintPreview"/> </remarks>
		/// <param name="enableChanges">optional object enableChanges</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintPreview(object enableChanges)
		{
			 Factory.ExecuteMethod(this, "PrintPreview", enableChanges);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.PrintPreview"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintPreview()
		{
			 Factory.ExecuteMethod(this, "PrintPreview");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Select"/> </remarks>
		/// <param name="replace">optional object replace</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Select(object replace)
		{
			 Factory.ExecuteMethod(this, "Select", replace);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.Select"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.PrintOut"/> </remarks>
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.PrintOut"/> </remarks>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="copies">optional object copies</param>
		/// <param name="preview">optional object preview</param>
		/// <param name="activePrinter">optional object activePrinter</param>
		/// <param name="printToFile">optional object printToFile</param>
		/// <param name="collate">optional object collate</param>
		/// <param name="prToFileName">optional object prToFileName</param>
		/// <param name="ignorePrintAreas">optional object ignorePrintAreas</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName, object ignorePrintAreas)
		{
			 Factory.ExecuteMethod(this, "PrintOut", new object[]{ from, to, copies, preview, activePrinter, printToFile, collate, prToFileName, ignorePrintAreas });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.PrintOut"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			 Factory.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.PrintOut"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.PrintOut"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.PrintOut"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.PrintOut"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.PrintOut"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.PrintOut"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Worksheets.PrintOut"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.worksheets.add2"/> </remarks>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		/// <param name="count">optional object count</param>
		/// <param name="newLayout">optional object newLayout</param>
		[SupportByVersion("Excel", 15, 16)]
		public object Add2(object before, object after, object count, object newLayout)
		{
			return Factory.ExecuteVariantMethodGet(this, "Add2", before, after, count, newLayout);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.worksheets.add2"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public object Add2()
		{
			return Factory.ExecuteVariantMethodGet(this, "Add2");
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.worksheets.add2"/> </remarks>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public object Add2(object before)
		{
			return Factory.ExecuteVariantMethodGet(this, "Add2", before);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.worksheets.add2"/> </remarks>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public object Add2(object before, object after)
		{
			return Factory.ExecuteVariantMethodGet(this, "Add2", before, after);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.worksheets.add2"/> </remarks>
		/// <param name="before">optional object before</param>
		/// <param name="after">optional object after</param>
		/// <param name="count">optional object count</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public object Add2(object before, object after, object count)
		{
			return Factory.ExecuteVariantMethodGet(this, "Add2", before, after, count);
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