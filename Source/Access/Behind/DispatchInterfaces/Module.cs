using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.AccessApi;

namespace NetOffice.AccessApi.Behind
{
	/// <summary>
	/// DispatchInterface Module 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835649.aspx </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Module : COMObject, NetOffice.AccessApi.Module
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
                    _contractType = typeof(NetOffice.AccessApi.Module);
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
                    _type = typeof(Module);                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Module() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197648.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Application>(this, "Application", typeof(NetOffice.AccessApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845790.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192086.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820960.aspx </remarks>
		/// <param name="line">Int32 line</param>
		/// <param name="numLines">Int32 numLines</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string get_Lines(Int32 line, Int32 numLines)
		{
			return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Lines", line, numLines);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Lines
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820960.aspx </remarks>
		/// <param name="line">Int32 line</param>
		/// <param name="numLines">Int32 numLines</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_Lines")]
		public virtual string Lines(Int32 line, Int32 numLines)
		{
			return get_Lines(line, numLines);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195500.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 CountOfLines
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CountOfLines");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836419.aspx </remarks>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 get_ProcStartLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ProcStartLine", procName, procKind);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ProcStartLine
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836419.aspx </remarks>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_ProcStartLine")]
		public virtual Int32 ProcStartLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return get_ProcStartLine(procName, procKind);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835086.aspx </remarks>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 get_ProcCountLines(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ProcCountLines", procName, procKind);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ProcCountLines
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835086.aspx </remarks>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_ProcCountLines")]
		public virtual Int32 ProcCountLines(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return get_ProcCountLines(procName, procKind);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822434.aspx </remarks>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 get_ProcBodyLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ProcBodyLine", procName, procKind);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ProcBodyLine
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822434.aspx </remarks>
		/// <param name="procName">string procName</param>
		/// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_ProcBodyLine")]
		public virtual Int32 ProcBodyLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
		{
			return get_ProcBodyLine(procName, procKind);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195085.aspx </remarks>
		/// <param name="line">Int32 line</param>
		/// <param name="pprockind">NetOffice.VBIDEApi.Enums.vbext_ProcKind pprockind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string get_ProcOfLine(Int32 line, out NetOffice.VBIDEApi.Enums.vbext_ProcKind pprockind)
		{		
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pprockind = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(line, pprockind);
			object returnItem = Invoker.PropertyGet(this, "ProcOfLine", paramsArray, modifiers);
			pprockind = (NetOffice.VBIDEApi.Enums.vbext_ProcKind)paramsArray[1];
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ProcOfLine
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195085.aspx </remarks>
		/// <param name="line">Int32 line</param>
		/// <param name="pprockind">NetOffice.VBIDEApi.Enums.vbext_ProcKind pprockind</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_ProcOfLine")]
		public virtual string ProcOfLine(Int32 line, out NetOffice.VBIDEApi.Enums.vbext_ProcKind pprockind)
		{
			return get_ProcOfLine(line, out pprockind);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837190.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 CountOfDeclarationLines
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CountOfDeclarationLines");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835633.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Enums.AcModuleType Type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcModuleType>(this, "Type");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845332.aspx </remarks>
		/// <param name="text">string text</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void InsertText(string text)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "InsertText", text);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845379.aspx </remarks>
		/// <param name="_string">string string</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void AddFromString(string _string)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddFromString", _string);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821352.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void AddFromFile(string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddFromFile", fileName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194137.aspx </remarks>
		/// <param name="line">Int32 line</param>
		/// <param name="_string">string string</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void InsertLines(Int32 line, string _string)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "InsertLines", line, _string);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194301.aspx </remarks>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="count">Int32 count</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void DeleteLines(Int32 startLine, Int32 count)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteLines", startLine, count);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198276.aspx </remarks>
		/// <param name="line">Int32 line</param>
		/// <param name="_string">string string</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void ReplaceLine(Int32 line, string _string)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ReplaceLine", line, _string);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845440.aspx </remarks>
		/// <param name="eventName">string eventName</param>
		/// <param name="objectName">string objectName</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 CreateEventProc(string eventName, string objectName)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CreateEventProc", eventName, objectName);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195471.aspx </remarks>
		/// <param name="target">string target</param>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="startColumn">Int32 startColumn</param>
		/// <param name="endLine">Int32 endLine</param>
		/// <param name="endColumn">Int32 endColumn</param>
		/// <param name="wholeWord">optional bool WholeWord = false</param>
		/// <param name="matchCase">optional bool MatchCase = false</param>
		/// <param name="patternSearch">optional bool PatternSearch = false</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord, object matchCase, object patternSearch)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Find", new object[]{ target, startLine, startColumn, endLine, endColumn, wholeWord, matchCase, patternSearch });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195471.aspx </remarks>
		/// <param name="target">string target</param>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="startColumn">Int32 startColumn</param>
		/// <param name="endLine">Int32 endLine</param>
		/// <param name="endColumn">Int32 endColumn</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Find", new object[]{ target, startLine, startColumn, endLine, endColumn });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195471.aspx </remarks>
		/// <param name="target">string target</param>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="startColumn">Int32 startColumn</param>
		/// <param name="endLine">Int32 endLine</param>
		/// <param name="endColumn">Int32 endColumn</param>
		/// <param name="wholeWord">optional bool WholeWord = false</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Find", new object[]{ target, startLine, startColumn, endLine, endColumn, wholeWord });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195471.aspx </remarks>
		/// <param name="target">string target</param>
		/// <param name="startLine">Int32 startLine</param>
		/// <param name="startColumn">Int32 startColumn</param>
		/// <param name="endLine">Int32 endLine</param>
		/// <param name="endColumn">Int32 endColumn</param>
		/// <param name="wholeWord">optional bool WholeWord = false</param>
		/// <param name="matchCase">optional bool MatchCase = false</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord, object matchCase)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Find", new object[]{ target, startLine, startColumn, endLine, endColumn, wholeWord, matchCase });
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public virtual bool IsMemberSafe(Int32 dispid)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsMemberSafe", dispid);
		}

		#endregion

		#pragma warning restore
	}
}


