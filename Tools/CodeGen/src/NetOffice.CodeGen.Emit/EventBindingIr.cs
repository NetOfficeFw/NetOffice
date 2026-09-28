using System;
using System.Collections.Generic;
using System.Linq;

namespace NetOffice.CodeGen.Emit
{
    /// <summary>One COM event connection-point sink owned by a co-class event bridge.</summary>
    public sealed class WrapperEventSinkBinding
    {
        public string SinkHelperType { get; set; }
        public string FieldName { get; set; }

        internal WrapperEventSinkBinding Canonicalize()
        {
            if (string.IsNullOrWhiteSpace(SinkHelperType)) throw new ArgumentException("WrapperEventSinkBinding.SinkHelperType is required.");
            if (string.IsNullOrWhiteSpace(FieldName)) throw new ArgumentException("WrapperEventSinkBinding.FieldName is required.");
            return new WrapperEventSinkBinding { SinkHelperType = SinkHelperType.Trim(), FieldName = FieldName.Trim() };
        }
    }

    /// <summary>Exact contract facets for emitter-owned IEventBinding members.</summary>
    public sealed class WrapperEventBindingMemberContracts
    {
        public WrapperRuntimeMemberContract CreateEventBridge { get; set; }
        public WrapperRuntimeMemberContract EventBridgeInitialized { get; set; }
        public WrapperRuntimeMemberContract HasEventRecipients { get; set; }
        public WrapperRuntimeMemberContract HasNamedEventRecipients { get; set; }
        public WrapperRuntimeMemberContract GetEventRecipients { get; set; }
        public WrapperRuntimeMemberContract GetCountOfEventRecipients { get; set; }
        public WrapperRuntimeMemberContract RaiseCustomEvent { get; set; }
        public WrapperRuntimeMemberContract DisposeEventBridge { get; set; }

        internal WrapperEventBindingMemberContracts Canonicalize()
        {
            return new WrapperEventBindingMemberContracts
            {
                CreateEventBridge = Canonicalize(CreateEventBridge),
                EventBridgeInitialized = Canonicalize(EventBridgeInitialized),
                HasEventRecipients = Canonicalize(HasEventRecipients),
                HasNamedEventRecipients = Canonicalize(HasNamedEventRecipients),
                GetEventRecipients = Canonicalize(GetEventRecipients),
                GetCountOfEventRecipients = Canonicalize(GetCountOfEventRecipients),
                RaiseCustomEvent = Canonicalize(RaiseCustomEvent),
                DisposeEventBridge = Canonicalize(DisposeEventBridge)
            };
        }

        internal IEnumerable<WrapperRuntimeMemberContract> All()
        {
            return new[] { CreateEventBridge, EventBridgeInitialized, HasEventRecipients, HasNamedEventRecipients, GetEventRecipients, GetCountOfEventRecipients, RaiseCustomEvent, DisposeEventBridge }.Where(value => value != null);
        }

        private static WrapperRuntimeMemberContract Canonicalize(WrapperRuntimeMemberContract value) => value == null ? null : value.Canonicalize();
    }

    /// <summary>Runtime plan implementing NetOffice.IEventBinding for a generated co-class.</summary>
    public sealed class WrapperEventBinding
    {
        public string ConnectPointType { get; set; } = "NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint";
        public string ConnectPointField { get; set; } = "_connectPoint";
        public string ActiveSinkIdField { get; set; } = "_activeSinkId";
        public string SinkHelperType { get; set; } = "SinkHelper";
        public string ReflectorType { get; set; } = "NetOffice.Events.CoClassEventReflector";
        public IList<WrapperEventSinkBinding> Sinks { get; set; } = new List<WrapperEventSinkBinding>();
        public WrapperEventBindingMemberContracts Contracts { get; set; }

        internal WrapperEventBinding Canonicalize()
        {
            var result = new WrapperEventBinding
            {
                ConnectPointType = Normalize(ConnectPointType, "NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint"),
                ConnectPointField = Normalize(ConnectPointField, "_connectPoint"),
                ActiveSinkIdField = Normalize(ActiveSinkIdField, "_activeSinkId"),
                SinkHelperType = Normalize(SinkHelperType, "SinkHelper"),
                ReflectorType = Normalize(ReflectorType, "NetOffice.Events.CoClassEventReflector"),
                Sinks = (Sinks ?? new List<WrapperEventSinkBinding>()).Select(value => value.Canonicalize()).ToList(),
                Contracts = Contracts == null ? null : Contracts.Canonicalize()
            };
            if (result.Sinks.Count == 0) throw new ArgumentException("WrapperEventBinding requires at least one sink.");
            if (result.Sinks.Select(value => value.FieldName).Distinct(StringComparer.Ordinal).Count() != result.Sinks.Count)
                throw new ArgumentException("WrapperEventBinding sink field names must be unique.");
            return result;
        }

        private static string Normalize(string value, string fallback) => string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
    }
}
