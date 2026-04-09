using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Diagnostics;
using Microsoft.CodeAnalysis.Operations;

namespace NetOffice.Analyzers;

[DiagnosticAnalyzer(LanguageNames.CSharp)]
public sealed class EventValidateRaiseConsistencyAnalyzer : DiagnosticAnalyzer
{
    public const string ValidateVsRaiseId = "NOE001";
    public const string RaiseVsMethodId = "NOE002";
    public const string CannotVerifyId = "NOE003";

    private static readonly DiagnosticDescriptor ValidateVsRaiseRule = new(
        id: ValidateVsRaiseId,
        title: "Validate event name must match RaiseCustomEvent event name",
        messageFormat: "Validate event name '{0}' does not match RaiseCustomEvent event name '{1}'",
        category: "NetOffice.Events",
        defaultSeverity: DiagnosticSeverity.Warning,
        isEnabledByDefault: true);

    private static readonly DiagnosticDescriptor RaiseVsMethodRule = new(
        id: RaiseVsMethodId,
        title: "RaiseCustomEvent event name should match method name",
        messageFormat: "RaiseCustomEvent event name '{0}' does not match containing method name '{1}'",
        category: "NetOffice.Events",
        defaultSeverity: DiagnosticSeverity.Info,
        isEnabledByDefault: true);

    private static readonly DiagnosticDescriptor CannotVerifyRule = new(
        id: CannotVerifyId,
        title: "Unable to verify event name consistency",
        messageFormat: "Unable to verify event name consistency in '{0}'. Use string literals for Validate/RaiseCustomEvent and avoid multiple distinct event names per method.",
        category: "NetOffice.Events",
        defaultSeverity: DiagnosticSeverity.Info,
        isEnabledByDefault: true);

    public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics =>
        ImmutableArray.Create(ValidateVsRaiseRule, RaiseVsMethodRule, CannotVerifyRule);

    public override void Initialize(AnalysisContext context)
    {
        context.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);
        context.EnableConcurrentExecution();

        context.RegisterOperationBlockStartAction(startContext =>
        {
            if (startContext.OwningSymbol is not IMethodSymbol methodSymbol)
                return;

            // This analyzer is meant for NetOffice sink helpers (derived from NetOffice.SinkHelper).
            // Scoping it keeps the signal high and avoids flagging unrelated "Validate"/"RaiseCustomEvent" methods.
            if (!InheritsFrom(methodSymbol.ContainingType, "NetOffice.SinkHelper"))
                return;

            var validateInvocations = new List<IInvocationOperation>();
            var raiseInvocations = new List<IInvocationOperation>();

            var validateConstNames = new HashSet<string>(StringComparer.Ordinal);
            var raiseConstNames = new HashSet<string>(StringComparer.Ordinal);

            bool validateHasNonConst = false;
            bool raiseHasNonConst = false;

            startContext.RegisterOperationAction(invocationContext =>
            {
                var invocation = (IInvocationOperation)invocationContext.Operation;
                var target = invocation.TargetMethod;

                if (IsValidateCall(target, invocation))
                {
                    validateInvocations.Add(invocation);

                    if (TryGetFirstStringConstant(invocation, out var validateName))
                        validateConstNames.Add(validateName);
                    else
                        validateHasNonConst = true;

                    return;
                }

                if (IsRaiseCustomEventCall(target, invocation))
                {
                    raiseInvocations.Add(invocation);

                    if (TryGetFirstStringConstant(invocation, out var raiseName))
                        raiseConstNames.Add(raiseName);
                    else
                        raiseHasNonConst = true;
                }
            }, OperationKind.Invocation);

            startContext.RegisterOperationBlockEndAction(endContext =>
            {
                if (validateInvocations.Count == 0 && raiseInvocations.Count == 0)
                    return;

                bool raisedAnyDiagnostic = false;

                // NOE001: Validate("...") must match RaiseCustomEvent("...") if both are verifiable.
                if (!validateHasNonConst &&
                    !raiseHasNonConst &&
                    validateConstNames.Count == 1 &&
                    raiseConstNames.Count == 1)
                {
                    string validateName = validateConstNames.Single();
                    string raiseName = raiseConstNames.Single();

                    if (!string.Equals(validateName, raiseName, StringComparison.Ordinal))
                    {
                        // Report on the Validate call so the fix is obvious (the bug class seen in DocumentEvents2.cs).
                        var location = validateInvocations[0].Syntax.GetLocation();
                        endContext.ReportDiagnostic(Diagnostic.Create(
                            ValidateVsRaiseRule,
                            location,
                            validateName,
                            raiseName));
                        raisedAnyDiagnostic = true;
                    }
                }

                // NOE002: Raised event name should match method name (convention in this codebase).
                if (!raiseHasNonConst && raiseConstNames.Count == 1)
                {
                    string raiseName = raiseConstNames.Single();
                    string methodName = methodSymbol.Name;

                    if (!string.Equals(raiseName, methodName, StringComparison.Ordinal))
                    {
                        var location = raiseInvocations[0].Syntax.GetLocation();
                        endContext.ReportDiagnostic(Diagnostic.Create(
                            RaiseVsMethodRule,
                            location,
                            raiseName,
                            methodName));
                        raisedAnyDiagnostic = true;
                    }
                }

                // NOE003: Analyzer cannot reliably check the method due to ambiguity / non-literal strings / missing pair.
                if (!raisedAnyDiagnostic && NeedsCannotVerify(validateInvocations, raiseInvocations, validateConstNames, raiseConstNames, validateHasNonConst, raiseHasNonConst))
                {
                    var location = GetMethodNameLocation(methodSymbol) ?? raiseInvocations.FirstOrDefault()?.Syntax.GetLocation() ?? validateInvocations.First().Syntax.GetLocation();
                    endContext.ReportDiagnostic(Diagnostic.Create(CannotVerifyRule, location, methodSymbol.Name));
                }
            });
        });
    }

    private static bool IsValidateCall(IMethodSymbol target, IInvocationOperation invocation)
    {
        if (!string.Equals(target.Name, "Validate", StringComparison.Ordinal))
            return false;

        if (invocation.Arguments.Length < 1)
            return false;

        var firstParam = invocation.Arguments[0].Parameter;
        if (firstParam == null || firstParam.Type.SpecialType != SpecialType.System_String)
            return false;

        // NetOffice uses SinkHelper.Validate(string) -> bool; this guards against unrelated Validate overloads.
        return target.ReturnType.SpecialType == SpecialType.System_Boolean;
    }

    private static bool IsRaiseCustomEventCall(IMethodSymbol target, IInvocationOperation invocation)
    {
        if (!string.Equals(target.Name, "RaiseCustomEvent", StringComparison.Ordinal))
            return false;

        if (invocation.Arguments.Length < 1)
            return false;

        var firstParam = invocation.Arguments[0].Parameter;
        if (firstParam == null || firstParam.Type.SpecialType != SpecialType.System_String)
            return false;

        // In NetOffice, EventBinding is an IEventBinding and RaiseCustomEvent returns int.
        return target.ReturnType.SpecialType == SpecialType.System_Int32;
    }

    private static bool TryGetFirstStringConstant(IInvocationOperation invocation, out string value)
    {
        value = "";
        if (invocation.Arguments.Length < 1)
            return false;

        var constant = invocation.Arguments[0].Value.ConstantValue;
        if (!constant.HasValue || constant.Value is not string s)
            return false;

        value = s;
        return true;
    }

    private static bool NeedsCannotVerify(
        List<IInvocationOperation> validateInvocations,
        List<IInvocationOperation> raiseInvocations,
        HashSet<string> validateConstNames,
        HashSet<string> raiseConstNames,
        bool validateHasNonConst,
        bool raiseHasNonConst)
    {
        // If the method is supposed to follow the NetOffice sink-helper pattern,
        // we need exactly one verifiable Validate name and one verifiable RaiseCustomEvent name.
        if (validateInvocations.Count == 0 || raiseInvocations.Count == 0)
            return true;

        if (validateHasNonConst || raiseHasNonConst)
            return true;

        if (validateConstNames.Count != 1 || raiseConstNames.Count != 1)
            return true;

        return false;
    }

    private static Location? GetMethodNameLocation(IMethodSymbol methodSymbol)
    {
        var syntaxRef = methodSymbol.DeclaringSyntaxReferences.FirstOrDefault();
        if (syntaxRef == null)
            return null;

        var node = syntaxRef.GetSyntax();
        if (node is MethodDeclarationSyntax methodDecl)
            return methodDecl.Identifier.GetLocation();

        // Best-effort: report on the declaration node.
        return node.GetLocation();
    }

    private static bool InheritsFrom(INamedTypeSymbol? type, string fullyQualifiedMetadataName)
    {
        for (var current = type; current != null; current = current.BaseType)
        {
            // e.g. "NetOffice.SinkHelper"
            var name = current.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat);
            if (name.StartsWith("global::", StringComparison.Ordinal))
                name = name.Substring("global::".Length);

            if (string.Equals(name, fullyQualifiedMetadataName, StringComparison.Ordinal))
                return true;
        }

        return false;
    }
}
