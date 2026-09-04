# NetOffice Core COM Lifetime Architecture

Status: Proposed

## Purpose

Make NetOffice Core manage Microsoft Office COM objects reliably without invalidating live managed references, releasing callback-owned objects, or depending on every caller to manually dispose an entire wrapper tree.

This specification targets the .NET Framework 4.6.2 runtime used by NetOffice. It preserves the generated Office API surface where practical, but deliberately changes unsafe lifetime behavior in the next major version. In this document, native Office objects are represented by runtime callable wrappers (RCWs). The generated managed event sink is exposed back to Office through a COM callable wrapper (CCW); RCW and CCW ownership are distinct.

## Background

Native COM lifetime is defined per interface pointer. An operation returning an interface pointer gives the caller a reference that must eventually be balanced by `IUnknown::Release`. `QueryInterface` also returns an independently releasable reference. The canonical `IUnknown` identity identifies the underlying COM object, but COM does not require other interface pointers for that object to have equal addresses.

The CLR interposes a runtime callable wrapper (RCW). Managed references point to the RCW, not directly to native interface pointers. The RCW caches interface pointers and releases its native references when the RCW is collected. A process normally has one RCW for a COM identity, so unrelated managed references can share the same RCW.

`Marshal.ReleaseComObject` decrements the RCW's internal COM-entry count. It does not release one arbitrary C# reference. Releasing the count to zero disconnects the shared RCW even when other managed code still has references to it. Releasing an RCW during an active call can cause access violations or process corruption. `Marshal.FinalReleaseComObject` is more dangerous because it forces the RCW count to zero regardless of other entries.

Normative external references:

- [Microsoft: Implementing Reference Counting](https://learn.microsoft.com/en-us/windows/win32/com/implementing-reference-counting)
- [Microsoft: Runtime Callable Wrapper](https://learn.microsoft.com/en-us/dotnet/standard/native-interop/runtime-callable-wrapper)
- [Microsoft: Marshal.ReleaseComObject](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.marshal.releasecomobject?view=netframework-4.8.1)

## Current architecture and failure modes

### Current ownership model

Each `COMObject` owns a `COMProxyShare`. The share contains an RCW and a NetOffice counter. Each constructed wrapper is also strongly retained by:

- its parent's `_listChildObjects`; and
- `Core._globalObjectList`.

Disposing a wrapper recursively disposes its children. When a share counter reaches zero, it invokes `Marshal.ReleaseComObject` on its RCW.

Relevant implementation:

- `Source/NetOffice/COMObject.cs`
- `Source/NetOffice/COMProxyShare.cs`
- `Source/NetOffice/Core.cs`
- `Source/NetOffice/Invoker.cs`
- generated `Source/*/Events/*.cs`

### F1: wrapper-local counters do not model RCW identity

`Core.CreateNewProxyShare` creates a new share for each wrapper construction. There is no registry keyed by canonical COM identity. Two wrappers around the same RCW can therefore have independent shares, each believing it owns the right to call `ReleaseComObject`.

This occurs naturally when Office returns the same singleton or object through different properties, when a callback supplies an object already held by application code, or when raw RCWs enter through public construction APIs.

### F2: strong tracking turns missed disposal into permanent retention

The parent child list and Core global list strongly reference every wrapper. A wrapper that becomes unreachable in application code does not become collectible while its parent or Core remains alive. Its RCW consequently remains rooted as well. The tracking intended to enable bulk cleanup therefore makes forgotten wrappers leak until an ancestor or all of Core is explicitly disposed.

### F3: recursive parent disposal invalidates live child references

A caller can retain a child in a field, then dispose an ancestor. Current recursive disposal disconnects the child's share even though the child remains reachable in C#. The next call through that child fails with disposed-wrapper or disconnected-RCW behavior.

Managed reachability cannot be inferred from the COM object-model path that produced a wrapper. Parentage is useful provenance and diagnostics, but it is not ownership.

### F4: callback parameters are borrowed, not owned by the sink

Generated event sinks call `Invoker.ReleaseParamsArray` when no handler is present. That method directly calls `Marshal.ReleaseComObject` for raw COM parameters. When handlers are present, callback arguments are wrapped as ordinary child objects; `EnableAutoDisposeEventArguments` can dispose them after delivery.

A callback parameter can map to an RCW already used elsewhere in the process. Direct release can therefore decrement or disconnect somebody else's shared RCW. Auto-disposal also makes it unsafe for a handler to retain a callback wrapper after returning.

### F5: raw RCWs can escape NetOffice accounting

`ICOMObjectProxy.UnderlyingObject` publicly exposes the RCW. Constructors also accept externally supplied RCWs. Core cannot know whether other code stores, invokes, or releases such an RCW. A NetOffice counter cannot prove exclusive RCW ownership after this boundary is crossed.

### F6: disposal is not coordinated with calls or COM apartments

The current counter protects its integer mutation, but there is no identity-level gate between invocation, callback delivery, event unsubscription, and release. One thread can release an RCW while another thread is invoking it.

Office automation objects are normally apartment-affine. `ReleaseComObject`, connection-point `Unadvise`, and final cleanup can require execution on the owning STA. A finalizer-thread release is not an acceptable fix.

### F7: wrapper replacement and cloning have ambiguous transfer semantics

Replacement constructors transfer an existing share without a formal move operation that invalidates the replaced wrapper. Clone paths first construct a wrapper around the RCW and then overwrite its newly created share with the original share. These paths make ownership counting dependent on constructor side effects and leave room for orphaned shares or double disposal.

## Required invariants

The implementation MUST maintain these invariants.

1. **Identity invariant:** All RCWs representing one canonical COM identity within one registry share one identity entry. The implementation MUST NOT claim coordination across AppDomain boundaries.
2. **Balanced identity probe:** Every `Marshal.GetIUnknownForObject` or `QueryInterface` used for identity discovery is balanced exactly once in `finally` with `Marshal.Release`.
3. **No manual RCW release:** NetOffice NEVER calls `Marshal.ReleaseComObject` or `Marshal.FinalReleaseComObject` for Office RCWs, callback parameters, externally supplied RCWs, connection points, or COM metadata RCWs. It drops managed references and lets the CLR release RCW-held native references.
4. **Single cleanup authority:** Only the identity registry may remove its strong RCW reference. Generated wrappers, sinks, enumerators, diagnostics, and helper collections MUST NOT perform RCW cleanup.
5. **Live wrapper safety:** Disposing one wrapper MUST NOT disconnect another live wrapper for the same COM identity.
6. **No ancestor ownership:** Disposing a parent MUST NOT dispose independently reachable child wrappers.
7. **Idempotent lease release:** A wrapper's lease can transition from active to released exactly once. Repeated `Dispose` and finalization are no-ops after the first transition.
8. **No in-flight cleanup:** Registry removal and connection teardown MUST NOT run while the identity has an active invocation or callback.
9. **Apartment affinity:** Invocation, `Advise`, `Unadvise`, and `Quit` MUST execute through the identity's owning apartment dispatcher.
10. **No strong diagnostic roots:** Diagnostics and parent-path tracking MUST NOT keep wrappers alive.
11. **Retainable callbacks:** A callback wrapper retained by user code remains usable after the callback returns.
12. **Escape visibility:** Access through a public raw-object API MUST be recorded because Core can no longer account for all managed aliases.
13. **Failure safety:** If apartment cleanup cannot run safely, Core MUST report the unresolved event connection and leave the RCW to CLR management; it MUST NOT release from an arbitrary thread.

## Proposed architecture

### 1. AppDomain-wide COM identity registry

Replace wrapper-local ownership with one `ComIdentityRegistry` shared by all `Core` instances in an AppDomain.

```csharp
internal sealed class ComIdentityRegistry
{
    ComLease Attach(object rcw, ComIngress ingress, ComApartment apartment);
    T Invoke<T>(ComLease lease, Func<object, T> operation);
    void MarkEscaped(ComIdentityEntry entry);
    void ReleaseLease(ComIdentityEntry entry);
}
```

The registry key is the canonical `IUnknown` pointer value obtained only for identity comparison:

```csharp
IntPtr identity = Marshal.GetIUnknownForObject(rcw);
try
{
    // Lookup or create the entry. Never retain this temporary AddRef.
}
finally
{
    Marshal.Release(identity);
}
```

The pointer value may be used as a key only while the entry strongly retains its RCW. An entry remains registered while calls, callbacks, or event teardown are pending. `Attach` atomically acquires a lease from the existing entry and cancels any not-yet-committed removal. The registry removes the key only after rechecking zero activity and atomically detaching the RCW reference; this prevents reattachment and pointer-reuse races.

Registry lookup and entry state changes MUST be serialized. No user callbacks or COM calls may run while holding the registry lock. A static registry is only AppDomain-wide on .NET Framework. Because Core cannot coordinate RCW use in another AppDomain or in external managed code, it has no authority to force RCW release.

### 2. Identity entry

```csharp
internal sealed class ComIdentityEntry
{
    object Rcw;
    ComApartment Apartment;
    int ActiveLeases;
    int ActiveCalls;
    int ActiveCallbacks;
    int ActiveConnections;
    ComIngressKinds SeenIngress;
    bool Escaped;
    bool RemovalPending;
}
```

An entry owns the only strong registry reference to the RCW. Wrapper objects reference the entry through leases rather than holding independent cleanup authority.

`ComIngressKinds` is diagnostic provenance, not release authority:

- `NetOfficeActivation`: RCW originated from a ProgID activation performed by NetOffice.
- `InvocationResult`: RCW entered through a generated late-bound call.
- `CallbackBorrowed`: RCW entered as a COM callback parameter.
- `ExternalBorrowed`: RCW entered through a public constructor or wrapping API.

An entry accumulates every observed ingress kind. Provenance never authorizes `ReleaseComObject`; it explains why an identity is shared and supports leak diagnostics.

### 3. Wrapper lease

Every `COMObject` owns one `ComLease`:

```csharp
internal sealed class ComLease
{
    ComIdentityEntry Entry;
    int State; // Active or Released

    T Invoke<T>(Func<object, T> operation);
    void Dispose();
    ~ComLease();
}
```

Creation increments `ActiveLeases`. `Dispose` atomically decrements it once. The finalizer performs the same managed state transition when application code forgets disposal.

The finalizer MUST NOT call COM or block on an apartment. It only changes managed counters and requests registry reevaluation. If the last lease is finalized, the registry first handles any event connections as specified below. Once leases, calls, callbacks, and connections are all zero, it atomically removes the identity key, clears `Rcw`, and lets normal CLR RCW finalization release native references.

Each wrapper construction receives a new lease, even if the identity entry already exists. Two separately returned wrappers therefore dispose independently. Multiple C# variables referencing the same wrapper instance retain normal `IDisposable` semantics: disposing that instance invalidates all aliases of that same instance.

### 4. Weak object graph and diagnostics

Replace `_globalObjectList` and `_listChildObjects` strong references with identity-safe weak records:

```csharp
internal sealed class WrapperRecord
{
    WeakReference Wrapper;
    WeakReference Parent;
    long Sequence;
    string WrapperType;
    string CreationSite;
}
```

Parent links remain available for diagnostics but confer no lifetime ownership. Core periodically removes dead records.

`ProxyCount` is redefined as the number of live wrapper leases, not the number of strongly rooted wrapper objects. Add separate counters for identities, borrowed ingresses, active calls, event connections, pending teardown, unresolved teardown, and escaped identities.

### 5. Explicit lifetime scopes

Introduce an optional `ComScope` with explicit ownership transfer:

```csharp
using (ComScope scope = core.CreateScope())
{
    Workbook book = scope.Track(application.Workbooks.Add());
    // Use book only inside this scope unless it is detached.
}
```

`Track` transfers the wrapper's lease to the scope. Disposing the scope releases that lease and invalidates wrappers still owned by the scope. A caller that intentionally retains a wrapper MUST call `scope.Detach(wrapper)` before disposal; detaching atomically transfers the lease back to the wrapper. Managed reachability is not used to guess whether a wrapper escaped.

Generated calls MAY automatically transfer returned wrappers to the current ambient scope. Ambient scope storage MUST be apartment-local and MUST NOT flow blindly across threads.

`DisposeChildInstances` is deprecated. In the next major version it releases only legacy scope-owned wrappers created through that parent; it does not dispose independently owned child wrapper instances. Recursive ancestor invalidation is removed from normal `Dispose`.

### 6. Invocation gate

All generated Core extension methods pass the complete COM operation into one gate:

```csharp
return lease.Invoke(rcw => invoker.Invoke(rcw, ...));
```

The gate:

1. rejects a released wrapper;
2. increments `ActiveCalls` under the entry lock;
3. dispatches the complete delegate to the owning apartment;
4. invokes COM without holding entry or registry locks;
5. decrements `ActiveCalls` in `finally`; and
6. reevaluates pending teardown or registry removal.

The gate MUST NOT return an RCW for later invocation. `Dispose` may mark a lease released while a call is active, but registry removal waits until `ActiveCalls == 0`. New calls through that released wrapper fail with `ObjectDisposedException`.

Shutdown first rejects new queue submissions, then drains work already accepted. An STA MUST NOT synchronously wait for work queued to itself. Reentrant callbacks and disposal must therefore change state and schedule teardown, never wait while inside the COM call.

### 7. Apartment dispatcher

A root object captures or is supplied an `IComApartmentDispatcher`. Child identities inherit it.

```csharp
public interface IComApartmentDispatcher
{
    bool HasAccess { get; }
    T Invoke<T>(Func<T> action);
    void Post(Action action);
}
```

Default construction requires an STA thread. For Windows Forms and WPF, Core may adapt the current `SynchronizationContext`. Headless automation uses an explicit NetOffice STA dispatcher with a pumped message loop.

Cross-apartment direct invocation is rejected unless a configured dispatcher marshals it. Cleanup posted after dispatcher shutdown is downgraded to CLR-managed cleanup. Core never blocks the finalizer thread waiting for an STA.

### 8. Callback borrowing and retention

Generated sink methods change as follows.

Current rejected-callback behavior:

```csharp
if (!Validate(eventName))
{
    Invoker.ReleaseParamsArray(rawArguments);
    return;
}
```

Required behavior:

```csharp
if (!Validate(eventName))
    return;
```

Raw callback RCWs are borrowed from the interop marshaler. Sinks MUST NOT release them.

For delivered events, `EventDispatchScope` attaches COM arguments with `ComIngress.CallbackBorrowed`, creates wrapper leases, and holds them strongly through dispatch. After dispatch it drops only its temporary references. It does not dispose wrappers. If a handler retained a wrapper, that wrapper remains alive and usable. Otherwise its lease is finalized and the borrowed identity entry drops its RCW naturally.

`EnableAutoDisposeEventArguments` is deprecated and ignored under the new lifetime engine. Its behavior is incompatible with retainable callbacks.

Ref/out primitive values continue to be copied back after handlers run. COM wrapper parameters are never released as part of ref/out propagation.

Callback dispatch increments `ActiveCallbacks` for every involved identity. Cleanup waits for all callback counters to reach zero.

### 9. Raw RCW escape boundary

Keep `UnderlyingObject` only for compatibility. Its public getter calls `MarkEscaped` before returning the RCW. Core-internal code uses a non-public call-scoped accessor.

After escape, diagnostics record that Core cannot account for every managed alias. Constructors and `COMObject.Create<T>(object)` likewise classify incoming RCWs as `ExternalBorrowed`. No escape or provenance state changes cleanup policy: explicit RCW release is always prohibited.

### 10. RCW cleanup policy

NetOffice uses CLR-managed RCW cleanup only:

1. wrapper and scope leases determine when Core still needs an RCW;
2. weak object tracking permits unreachable wrappers to finalize;
3. the identity registry clears its strong RCW reference after leases, calls, callbacks, and event connections reach zero;
4. CLR garbage collection and RCW finalization release the native references held by that RCW.

Core MUST NOT call `ReleaseComObject`, even for an RCW created by NetOffice. .NET Framework can share an RCW across Core instances, external code, apartments, or AppDomains that Core cannot observe. A local exclusivity counter therefore cannot prove that forced release is safe.

This policy gives safety and eventual cleanup, not a promise that a native Office reference is released at the instant `Dispose` returns. `Quit`, explicit wrapper/scope disposal, and releasing Core's strong references remain deterministic operations. Native destruction timing remains controlled by the CLR.

If strict process termination is required regardless of third-party RCW aliases, automation MUST run in a dedicated broker process. Broker shutdown may terminate that process after graceful `Quit`; in-process `ReleaseComObject` is not an equivalent isolation boundary.

### 11. Events and connection points

`Advise` causes the COM server to retain the managed sink through its CCW. NetOffice must strongly retain the sink and connection record until `Unadvise` completes.

Explicit event source disposal:

1. marks the wrapper lease released;
2. prevents new managed event delivery;
3. posts `Unadvise` to the owning apartment;
4. retains the connection record, sink, and identity entry until completion;
5. decrements `ActiveConnections` in `finally`; and
6. reevaluates registry removal.

If the last source lease is finalized while connections remain, the finalizer schedules the same apartment-affine `Unadvise`; it never performs COM work itself. A new `Attach` before teardown commits cancels pending removal but does not silently discard a required `Unadvise`.

If the apartment dispatcher has already shut down, Core records an unresolved connection teardown failure. It drops no connection record required to keep the CCW valid and makes no arbitrary-thread COM call. `OfficeSession` must prevent this state by draining event teardown before dispatcher shutdown.

`SinkHelper.DisposeAll` must snapshot active sinks without holding a global lock during COM calls. Sink registries use weak source references and explicit strong connection records.

### 12. Root shutdown and Quit

`Application.Dispose` does not implicitly invalidate descendants. `Quit` is a server operation distinct from RCW release.

Provide an explicit session API:

```csharp
using (OfficeSession session = OfficeSession.StartExcel())
{
    Application application = session.Application;
}
```

Session shutdown:

1. stops accepting new work;
2. schedules all event unsubscriptions on the owning STA;
3. drains already accepted calls and connection teardown without an STA self-wait;
4. invokes `Quit` when configured;
5. releases session-owned scope leases;
6. reports still-live wrapper leases and unresolved connections without invalidating independently owned wrappers;
7. drops zero-activity registry RCW references; and
8. shuts down the dispatcher only after the queue is drained.

If wrappers remain live, the session reports them and leaves their RCWs CLR-managed. Safety takes precedence over forcing Excel to disappear.

## API changes

### New

- `ComScope`
- `OfficeSession`
- `IComApartmentDispatcher`
- `Core.CreateScope()`
- `Core.LifetimeDiagnostics`
- `ComIngressKinds` as internal diagnostic provenance
- structured events: identity attached, RCW escaped, teardown scheduled, teardown failed, registry reference dropped

### Changed

- `COMObject.Dispose()` releases only that wrapper's lease and its own event connection.
- Parent and global tracking become weak.
- `ProxyCount` counts active leases.
- `UnderlyingObject` records that the RCW escaped Core accounting.
- Event arguments are retainable and are not auto-disposed.
- `DisposeAllCOMProxies()` releases Core/scope leases but does not invalidate independently live wrappers.

### Deprecated

- `DisposeChildInstances`
- `EnableAutoDisposeEventArguments`
- direct public use of `Invoker.ReleaseParam` and `ReleaseParamsArray`
- custom `COMProxyShare` replacement through `Core.CreateProxyShare`

### Removed from Core internals

- every direct `Marshal.ReleaseComObject` and `Marshal.FinalReleaseComObject` call
- wrapper-local RCW cleanup authority
- recursive child disposal
- strong global wrapper ownership
- clone paths that construct and then overwrite a share

## Architecture review disposition

An independent advisor review rejected the first draft and identified five correctness defects. This revision incorporates all five:

1. Per-Core `ExclusiveOwned` release could disconnect an RCW shared with another Core or external code. Resolution: AppDomain-wide coalescing and CLR-managed RCW cleanup only.
2. Dispatching only call entry did not move the COM operation onto the STA. Resolution: the invocation gate accepts and dispatches the complete operation delegate.
3. A scope that added a second lease could not provide deterministic scope semantics. Resolution: `Track` transfers ownership and `Detach` explicitly transfers it back.
4. A finalized subscribed source could leave `ActiveConnections` permanently nonzero. Resolution: last-source finalization posts `Unadvise` while strong connection records keep the CCW valid.
5. Removing an identity key before pending cleanup allowed reattachment races. Resolution: entries remain registered until activity is rechecked and RCW detachment commits atomically.

Review disposition: accepted after these corrections. The recommended design intentionally omits manual RCW release.

## Migration strategy

### Phase 1: safety boundary

1. Stop releasing raw callback arguments.
2. Remove callback auto-disposal under an opt-in `SafeLifetimeManagement` setting.
3. Route invocations through the complete-delegate activity gate.
4. Add diagnostics for raw RCW escape and remaining direct release sites.
5. Replace direct releases of temporary COM metadata and connection objects with reference dropping.

This phase can be shipped compatibly behind a setting except where a direct release is proven capable of disconnecting a shared RCW.

### Phase 2: identity registry

1. Introduce the AppDomain-wide `ComIdentityRegistry`, entries, and leases shared by every Core.
2. Make generated wrappers attach through the registry.
3. Rewrite clone and replacement as explicit lease acquisition or transfer.
4. Convert parent and global lists to weak records.
5. Implement finalizer-triggered registry reevaluation without finalizer-thread COM calls.
6. Keep entries registered through pending event teardown and removal.

### Phase 3: major-version semantics

1. Change `Dispose` and `DisposeChildInstances` semantics.
2. Add ownership-transferring `ComScope` and `OfficeSession`.
3. Remove callback auto-disposal.
4. Require or create an STA dispatcher for root automation.
5. Update generated event sinks and collection enumerators.
6. Remove all `ReleaseComObject` and `FinalReleaseComObject` calls from Core and generated libraries.

## Verification specification

Tests MUST exercise observable lifetime behavior, not private counters alone.

### Managed model tests

- Two wrappers for one identity dispose independently.
- Repeated disposal and finalization release one lease exactly once.
- Parent disposal leaves a retained child callable.
- Weak parent and global records do not prevent wrapper collection.
- Scope disposal invalidates scope-owned wrappers.
- A wrapper detached before scope disposal remains callable.
- Replacement transfers one lease and invalidates only the moved-from wrapper.
- Clone acquires one additional lease without constructing an orphan share.
- Registry reference removal waits for active calls, callbacks, and event connections.
- Raw RCW access records an escape without changing safe cleanup behavior.
- Finalization of the last subscribed source schedules `Unadvise`.
- Reattachment while removal is pending reuses the registered identity entry.

### Windows COM integration tests

Use an in-process test COM server that records `AddRef`, `Release`, interface identity, thread ID, and callback nesting. Tests may force full collection only after proving all NetOffice strong references were dropped. They MUST NOT call `ReleaseComObject` to make an assertion pass.

Required scenarios:

1. Same `IUnknown` returned through different interfaces and repeated calls.
2. Same RCW returned to two independent wrappers.
3. Same object already held by a wrapper arrives in a callback.
4. Callback with no subscriber performs no NetOffice release.
5. Handler retains callback argument and calls it after callback completion.
6. Handler does not retain callback argument; wrapper becomes collectible.
7. Dispose races with a long-running call; registry removal occurs only after the call returns.
8. Reentrant callback disposes the event source without deadlock or teardown during dispatch.
9. Event `Advise`, `Unadvise`, invocation, and `Quit` execute on the owning STA.
10. Finalization of a subscribed source posts `Unadvise` and keeps its CCW alive through completion.
11. Dispatcher shutdown reports unresolved connection teardown instead of releasing from an arbitrary thread.
12. Public raw RCW escape is recorded and the alias remains functional after NetOffice wrapper disposal.
13. External code shares the RCW and remains functional after NetOffice wrapper disposal.

### Office smoke matrix

Exercise Excel, Word, Outlook, PowerPoint, and Access where installed:

- x86 Office with x86 process;
- x64 Office with x64 process;
- .NET Framework 4.6.2 and 4.8 runtime environments;
- automation application and COM add-in contexts;
- event-heavy and enumeration-heavy object graphs.

Observe process exit only after explicit `Quit`; do not treat forced process exit as proof of correct reference accounting.

## Acceptance criteria

The architecture is complete when:

- generated code contains no direct release of callback parameters;
- Core and generated libraries contain no `ReleaseComObject` or `FinalReleaseComObject` calls;
- forgotten wrappers are collectible despite a live Core or parent;
- disposing an ancestor cannot invalidate a retained descendant;
- disposing one wrapper cannot disconnect another wrapper for the same identity;
- retained callback wrappers remain usable;
- registry reference removal and connection teardown never overlap an invocation or callback;
- `Advise`, `Unadvise`, invocation, and `Quit` occur on the owning apartment;
- orphaned event connections are unadvised before dispatcher shutdown or reported as unresolved;
- diagnostics identify live leases, ingress provenance, and RCW escapes without rooting wrappers; and
- the Windows COM integration server confirms that NetOffice adds no unbalanced native references and performs no forced RCW release.

## Non-goals

- Reimplementing the CLR RCW.
- Calling native `AddRef` to manufacture ownership outside CLR accounting.
- Guaranteeing Office process termination while application code still holds live wrappers or escaped RCWs.
- Hiding cross-thread Office automation by silently creating arbitrary proxies.
- Using `ReleaseComObject` or `FinalReleaseComObject` as a cleanup shortcut.
