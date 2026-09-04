# NetOffice Core COM Lifetime Test Harness Architecture

Status: Proposed

## Purpose

Define a deterministic Windows integration harness that exercises NetOffice against instrumented native C++ COM objects and verifies the lifetime contract in [NetOffice Core COM Lifetime Architecture](Core-COM-Lifetime-Architecture.md).

The harness must prove consumer-visible safety and native lifetime invariants. It must not infer correctness from CLR runtime callable wrapper (RCW) internals, fixed `AddRef`/`Release` sequences, Office process exit, or successful test-process termination.

## Terminology

- **Native object:** a C++ COM implementation owned by the fixture.
- **RCW:** the CLR runtime callable wrapper representing a native COM identity in managed code.
- **NetOffice wrapper:** a generated or Core `COMObject` façade holding a NetOffice lease over an RCW.
- **Early-bound C# wrapper:** a C# `[ComImport]` interface or class from the fixture type library. It is another managed view of an RCW, not a second native object.
- **Raw alias:** an `object`, `dynamic`, or `[ComImport]` reference to an RCW outside a NetOffice wrapper.
- **CCW:** the COM callable wrapper used when the native fixture calls a managed event sink or managed callback implementation.
- **Canonical identity:** the controlling `IUnknown` obtained by `QueryInterface(IID_IUnknown)`.
- **Quiescent epoch:** a bounded interval in which managed work is drained, no test code invokes COM, the owning STA is pumped, and lifetime observations can be evaluated without deliberate new references.
- **Fixture reference count:** the native implementation's diagnostic count after an `AddRef` or `Release`. It is evidence about that fixture object, not the CLR's private RCW-entry count.

## Goals

The harness must:

1. exercise actual generated NetOffice wrappers, generated event sinks, `Invoker`, identity registration, lease disposal, scopes, and apartment dispatch;
2. observe every native fixture object's construction, interface negotiation, reference operations, calls, callbacks, event connections, and destruction;
3. verify that one NetOffice wrapper can be disposed or finalized without unsafe COM cleanup;
4. verify that independent wrappers over one canonical identity do not invalidate one another;
5. verify callback argument and event-sink lifetime across delivery, retention, reentrancy, exceptions, and teardown;
6. verify coexistence with early-bound C# interfaces, raw RCWs, `dynamic`, and managed CCW callbacks;
7. cover returned objects, tear-off interfaces, aggregation, enumerators, errors, scopes, parent/child graphs, concurrency, apartments, and shutdown;
8. isolate access violations, hangs, and corrupted RCWs from the test controller; and
9. produce a replayable trace that explains every failure without inspecting the object through COM.

## Non-goals

- Reimplement or inspect CLR RCW bookkeeping.
- Require a fixed number or ordering of incidental CLR `AddRef`, `Release`, or `QueryInterface` operations.
- Use `Marshal.ReleaseComObject` or `Marshal.FinalReleaseComObject` in production code or test cleanup.
- Treat `Dispose` as a promise of immediate native destruction.
- Treat COM server or Office process exit as proof of balanced references.
- Prove coordination across AppDomains, processes, or code that forcibly releases a shared RCW.
- Replace real Office compatibility smoke tests.

## Required architecture

### Process topology

```text
LifetimeTest.Controller
  |-- Job Object + deadlines + dump capture
  |-- native telemetry reader
  |-- managed result reader
  |
  +-- LifetimeTest.ScenarioHost.x86 (.NET Framework)
  |     |-- pumped owner STA
  |     |-- optional second STA and MTA workers
  |     |-- NetOffice Core under test
  |     |-- LifetimeFixture.NetOfficeApi
  |     |-- LifetimeFixture.Interop
  |     +-- NativeLifetimeFixture.x86.dll
  |
  +-- LifetimeTest.ScenarioHost.x64 (.NET Framework)
        |-- same managed layers
        +-- NativeLifetimeFixture.x64.dll
```

Every native integration scenario runs in a fresh child process. RCW caches, static identity registries, finalizers, connection points, pointer reuse, and leaked native state must not contaminate another scenario.

The controller places each host and its descendants in a Windows Job Object. It owns separate startup, progress-barrier, and total deadlines. On a hang or crash it captures a dump and terminates the entire Job Object. A failing COM scenario must not terminate or poison the test runner.

### Project layout

The implementation should use these logical projects:

```text
Tests/ComLifetime/
  NativeLifetimeFixture/          C++ DLL, IDL, class factories, telemetry
  LifetimeFixture.Interop/        C# ComImport types generated from the TLB
  LifetimeFixture.NetOfficeApi/   generated NetOffice-style façade and sinks
  LifetimeTest.Protocol/          non-COM trace/result schema
  LifetimeTest.ScenarioHost/      architecture-specific .NET Framework host
  LifetimeTest.Controller/        process orchestration and offline oracles
  LifetimeTest.HarnessTests/      fixture and oracle qualification tests
```

The fixture API must be generated through the same generator and templates used by production NetOffice assemblies. Hand-constructing `COMObject` instances is useful for Core model tests, but cannot prove generated property access, enumerators, or event sinks obey the lifetime contract.

### Build and activation

Build the native fixture, scenario host, and all native dependencies for both `x86` and `x64`. The controller rejects a row when managed and native architectures differ; `AnyCPU` is not permitted for the scenario host.

Use registration-free COM with a process activation context and side-by-side manifest. The manifest declares fixture CLSIDs, ProgIDs, threading models, and type library. Activation tests must cover both CLSID activation and the ProgID path used by generated NetOffice roots. The harness must not call `regsvr32`, mutate machine or per-user COM registration, or use `CoRegisterClassObject` as an activation substitute. This keeps `windows-2025` `dotnet test` execution non-administrative and prevents parallel or stale registration from selecting the wrong binary.

The native binary exports normal COM class factories. Class-factory and DLL module-lock counts are traced separately and excluded from object-lifetime assertions.

## Native fixture object model

### IDL requirements

The fixture type library must include:

- dual `IDispatch` automation interfaces used by late-bound NetOffice invocation;
- early-bound vtable interfaces consumed by C# `[ComImport]` types;
- methods and properties returning `IUnknown`, `IDispatch`, specific interfaces, `VARIANT`, `SAFEARRAY`, and `IEnumVARIANT`;
- optional arguments and `VARIANT_BOOL` values;
- `ref` and `out` scalar, `VARIANT`, and interface parameters;
- a source dispinterface and `IConnectionPointContainer`;
- a managed callback interface implemented by a C# CCW; and
- deterministic barrier methods for calls, callbacks, and teardown.

C++ exceptions must never cross the COM boundary. Failure objects return controlled HRESULTs and populate `IErrorInfo`.

### Identity family

`Root` is the primary automation object. It exposes `IRoot`, `IAlpha`, `IBeta`, `IDispatch`, and `IConnectionPointContainer`.

- `IAlpha` and `IBeta` use distinct interface-pointer tokens.
- `QueryInterface(IID_IUnknown)` from either returns the same controlling unknown.
- `GetSelf`, `Self` property access, and selected collection items repeatedly return `Root`.
- `GetNewChild` returns a new canonical identity on every call.
- `GetSharedChild` returns one cached child identity.
- `GetSibling` returns a distinct identity with similar behavior.
- `GetNull` returns a null interface pointer.

`TearOff` exposes a separately allocated interface implementation whose controlling `IUnknown` remains `Root`. It detects registries that key on an arbitrary interface pointer instead of canonical identity.

`AggregatedRoot` contains an aggregatable `Inner`. Its internal nondelegating unknown is never exposed. Every external `Inner` interface delegates `IUnknown` to the outer controlling unknown. `SeparateSibling` has a different controlling unknown. These objects distinguish legal COM aggregation from ordinary parent/child relationships.

### Graph and collection family

`Child` can return its parent, a shared sibling, and a new grandchild. This creates cycles and repeated paths without changing canonical ownership.

`ProbeCollection` implements index access and `_NewEnum`. Its `IEnumVARIANT` supports:

- scalar elements;
- `VT_UNKNOWN` and `VT_DISPATCH` views of the same `Root`;
- repeated identities;
- distinct child identities;
- `Next`, `Skip`, `Reset`, and `Clone`;
- early consumer termination;
- a controlled failure after a configured item; and
- abandonment without explicit enumerator disposal.

Array methods return `SAFEARRAY` values containing mixed scalars, nulls, repeated COM identities, and distinct objects.

### Invocation and race family

`BlockingObject.LongCall` signals `CallEntered`, waits on `AllowCallExit`, then signals `CallExited`. It retains itself for the native call according to normal C++ method execution rules.

`FailureObject` provides:

- `FailBeforeResult`;
- `ReturnObjectThenFail` where marshaling cleanup is observable;
- invalid argument and type-mismatch errors;
- `IErrorInfo` text and source; and
- a successful call after each failure.

`ReentrantObject` invokes a managed callback while a native call is active. It supports callback-initiated calls, wrapper disposal, event removal, and shutdown attempts.

### Connection-point and callback family

`ProbeConnectionPoint` follows COM connection-point ownership:

1. successful `Advise` obtains and retains exactly one server-owned sink reference associated with a unique cookie;
2. callback delivery temporarily protects the sink according to the implementation's documented rule;
3. successful `Unadvise` removes the cookie and releases the server-owned sink reference exactly once;
4. duplicate or unknown-cookie `Unadvise` returns the configured HRESULT without releasing another reference; and
5. controlled `Unadvise` blocking and failure are available.

The source can fire:

- no-argument notifications;
- a native object already held elsewhere in managed code;
- a newly created object;
- the same object through two interface types;
- `ref` and `out` primitives;
- `ref` and `out` COM interfaces, including substitution with `Sibling` or null;
- nested and recursive callbacks;
- a callback on the owner STA;
- a callback from an MTA worker requiring COM marshaling; and
- a callback after a controlled source-side delay.

`CallbackDriver` accepts an `IManagedProbeCallback` implemented in C#. This explicitly tests a user CCW independently of the generated NetOffice event-sink CCW. It can retain the callback across calls, release it, invoke it reentrantly, and invoke it from a worker apartment.

### Lifetime implementation rules

Every fixture object uses a thread-safe reference count and a stable monotonically allocated object ID. The object ID is never a raw address. Destruction publishes a tombstone record to telemetry before freeing memory. The telemetry store retains facts, never the COM object or an interface pointer.

The fixture must fail fast in an isolated host on reference-count underflow, duplicate destruction, use after destruction detected before free, invalid connection-cookie release, or apartment-contract violation. A raw stale-pointer access may instead produce an access violation; the controller classifies that as a scenario failure and preserves the dump.

## Ownership-neutral telemetry

### Transport

Native observations use a preallocated, versioned memory-mapped ring buffer created by the controller and opened by the fixture with a run-unique name and nonce. The fixture writes fixed-size append-only records using an atomic sequence. The controller reads records without entering COM.

The ring buffer must not allocate, query, `AddRef`, or retain the object being observed. Overflow, a torn record, schema mismatch, unknown run nonce, or non-monotonic committed sequence fails the scenario.

Managed diagnostics use a separate named pipe or append-only file. They must contain identifiers and immutable metadata only. A diagnostic record must never retain a wrapper, RCW, delegate, exception, or event source.

### Native trace schema

Every record contains:

- schema version, record size, run ID, scenario ID, and epoch ID;
- global sequence and `QueryPerformanceCounter` timestamp;
- process ID, thread ID, and `CoGetApartmentType` result;
- object ID, controlling-object ID, object kind, interface IID, and opaque pointer token;
- operation and phase;
- fixture reference count after `AddRef` or `Release`, where meaningful;
- HRESULT;
- native call depth and callback depth;
- connection cookie and sink token;
- barrier ID and state; and
- optional correlation ID supplied by the scenario host.

Native operations include factory create/release, construct, `QueryInterface` request/result, `AddRef`, `Release`, method enter/exit, enumerator operations, `Advise`, `Unadvise`, server-held sink `AddRef`/`Release`, callback enter/exit, barrier transitions, destroy, and invariant failure.

### Managed trace schema

Once the proposed Core diagnostics exist, every record contains:

- run, scenario, epoch, and correlation IDs;
- wrapper ID, identity-entry ID, and lease ID;
- wrapper type and ingress kind;
- dispatcher ID, source thread, destination thread, and apartment;
- lease, active-call, active-callback, and active-connection transitions;
- raw-RCW escape marking;
- removal requested, cancelled, committed, and registry-reference-dropped states;
- event teardown scheduled, started, completed, failed, or unresolved;
- weak-wrapper collection; and
- structured failure details without strong object references.

Trace assertions complement consumer behavior. Private counters alone are never acceptance oracles.

## Deterministic execution protocol

### Scenario lifecycle

For each scenario the controller:

1. creates the Job Object, trace mapping, result path, nonce, and deadlines;
2. launches an architecture-matched host with one scenario ID and deterministic seed;
3. waits for activation and owner-STA readiness barriers;
4. sends the barrier schedule when the scenario requires interleaving;
5. consumes heartbeat and progress records without invoking COM;
6. waits for the host result and process exit;
7. captures a dump before termination on crash or deadline failure;
8. terminates remaining Job Object processes;
9. evaluates native trace, managed trace, and result manifest offline; and
10. retains the seed, schedule, traces, stdout, stderr, and dump on failure.

Exit code zero is necessary but insufficient. Missing teardown records, trace overflow, an unconsumed expected barrier, or a violated oracle still fails the row.

### STA and MTA discipline

The host creates a dedicated owner STA with a continuously pumped Windows message loop. Fixture activation, generated-wrapper calls, `Advise`, `Unadvise`, and session shutdown execute through the tested dispatcher.

MTA workers call `CoInitializeEx` through CLR-compatible thread setup and identify their apartment in trace. A second independently pumped STA is available for cross-apartment cases. Tests must verify the thread and apartment at native method entry; observing a COM proxy on the caller does not prove owner-STA execution.

The controller never waits for work from inside the STA. Host waits pump messages when COM reentrancy requires it. Sleep-based synchronization is forbidden.

### Barriers and schedules

Named barriers include:

- `CallEntered` and `AllowCallExit`;
- `CallbackEntered` and `AllowCallbackExit`;
- `UnadviseEntered` and `AllowUnadviseExit`;
- `RemovalPending` and `AllowRemovalCommit`;
- `DestroyObserved`; and
- `DispatcherDrainStarted` and `DispatcherStopped`.

Each race has a small explicit schedule set. The full trace must prove the required happens-before relation. Random stress is supplemental and records its seed; it never substitutes for deterministic schedules.

### GC and finalizer protocol

GC-sensitive setup and release run in `[MethodImpl(MethodImplOptions.NoInlining | MethodImplOptions.NoOptimization)]` helpers. A release helper returns only weak references and immutable IDs. It must not return delegates, exceptions, assertion closures, or result objects that capture wrappers.

Aliases intended to survive use `GC.KeepAlive(alias)` after the final asserted call. Objects intended for collection leave their helper frame before the controller requests quiescence.

A bounded collection round is:

1. drain accepted dispatcher work;
2. run `GC.Collect()`;
3. run `GC.WaitForPendingFinalizers()` without blocking the owner STA;
4. pump and drain teardown posted to the owner STA;
5. run `GC.Collect()` again; and
6. inspect weak references and out-of-band telemetry.

The round may repeat to a documented maximum. Failure to collect is a leak signal only after diagnostics prove no intentional strong NetOffice reference, event connection, native callback hold, or test local remains. Tests never call a manual RCW-release API to accelerate collection.

## Oracle model

### Facts the harness may assert

- canonical `IUnknown` equality or inequality, with every explicit identity probe balanced in `finally` by `Marshal.Release`;
- success or documented managed failure of a call through a live or disposed wrapper;
- native object ID returned by independently held aliases;
- wrapper, lease, connection, and ingress transitions exposed by Core diagnostics;
- call, callback, teardown, and destruction ordering from trace;
- native reference-count non-underflow and exactly one destruction tombstone;
- eventual destruction after all managed and native holds are proven absent;
- server-owned sink reference balance from successful `Advise` through successful `Unadvise`;
- thread and apartment at native entry; and
- absence of forbidden managed calls from an assembly scan.

### Facts the harness must not assert

- an exact RCW internal entry count;
- one RCW per identity across different AppDomains, processes, or CLR instances;
- a fixed number or global ordering of native `AddRef`, `Release`, or `QueryInterface` calls caused by CLR activation, interface caching, marshaling, reflection, debugging, or finalization;
- immediate native destruction after `COMObject.Dispose`;
- raw pointer equality between different interfaces;
- process exit as lifetime balance; or
- attribution of an arbitrary native `Release` to NetOffice rather than CLR marshaling or finalization.

### Bounded native-balance assertions

Native `AddRef` and `Release` records are facts. Their safe uses are:

1. the fixture count never underflows;
2. a fixture identity reaches one destruction event after every documented native and managed hold is removed;
3. an event sink hold added by the fixture is released once for its cookie;
4. an explicit test-owned `GetIUnknownForObject` or `QueryInterface` probe has net-zero effect inside a fenced quiescent epoch; and
5. fixture self-tests balance known direct native operations.

The test suite must not compare a whole NetOffice scenario to one golden sequence. CLR versions, JIT mode, apartment marshaling, and interface caching legitimately change incidental sequences.

### Forbidden-call oracle

A build-time analyzer scans IL in NetOffice Core and every generated fixture wrapper/sink for calls to:

- `Marshal.ReleaseComObject`; and
- `Marshal.FinalReleaseComObject`.

Any call is a failure, including unreachable or helper code. The analyzer itself is qualified against a small assembly containing both forbidden calls. This static oracle is required because native telemetry cannot attribute every `Release` to its managed origin.

## Required scenario matrix

### A. Single NetOffice wrapper

| ID | Scenario | Required oracle |
| --- | --- | --- |
| A01 | Activate `Root`, invoke property and method, dispose once | Calls return the stable object ID; disposal releases one lease; no forced-release API exists. |
| A02 | Dispose the same wrapper twice | One lease transition; the second call is inert; native trace has no cleanup attributed to the second disposal. |
| A03 | Invoke after disposal | `ObjectDisposedException` before native method entry. |
| A04 | Forget a wrapper while `Core` remains alive | Wrapper becomes weakly collectible; weak diagnostics do not root it. |
| A05 | Finalize a wrapper during an unrelated live identity | Only its lease changes; the unrelated identity remains callable. |
| A06 | Throw a COM error, then reuse and dispose the wrapper | HRESULT and `IErrorInfo` are preserved; active-call state unwinds; the next call succeeds. |

### B. Multiple wrappers over one native identity

| ID | Scenario | Required oracle |
| --- | --- | --- |
| B01 | Wrap the same RCW twice, dispose wrapper A | Wrapper B remains callable and returns the same object ID. |
| B02 | Repeated property and method ingress returns `Root` 100 times | All wrappers map to one identity entry; disposing any 99 leaves the last callable. |
| B03 | Receive `IAlpha` and distinct-pointer `IBeta` tear-off | Interface tokens differ, canonical `IUnknown` agrees, and wrappers coordinate one identity. |
| B04 | Receive externally aggregated `Inner` and `AggregatedRoot` | Both map to the outer identity; `SeparateSibling` does not coalesce. |
| B05 | Dispose wrappers in every order, including finalization of one | Every wrapper releases one lease at most once; remaining wrappers stay usable. |
| B06 | Alias the same NetOffice wrapper in two C# variables | Disposing the instance invalidates both aliases because they are one wrapper instance. |
| B07 | Reattach while registry removal is pending | The registered entry is reused and pending removal is cancelled before commit. |
| B08 | Native allocator later reuses an address after destruction | A new identity epoch is not confused with the tombstoned object ID or old registry entry. |

### C. Callback and event lifecycle

| ID | Scenario | Required oracle |
| --- | --- | --- |
| C01 | Callback passes `Root` already held by a normal wrapper | Callback ingress is borrowed and coalesces with the existing identity; both wrappers remain independently usable. |
| C02 | Fire an object callback with no managed subscriber | Callback returns normally; a pre-held alias remains callable; no argument release/disposal is emitted by NetOffice. |
| C03 | Handler retains callback wrapper after return | A later call succeeds with the same object ID; dropping it later permits bounded collection. |
| C04 | Handler does not retain callback wrapper | It survives through callback exit, then becomes weakly collectible without destroying another held alias. |
| C05 | Handler mutates `ref`/`out` primitives and substitutes a COM interface | Native caller receives exact values and substituted canonical identity; original and substitute remain valid according to their holds. |
| C06 | Handler reenters source and disposes its source wrapper | Reentrant call completes; no deadlock; `Unadvise` starts only after callback exit. |
| C07 | Handler throws | Documented callback HRESULT is returned; callback/activity counters unwind; later event or teardown proceeds. |
| C08 | Subscribe, fire, unsubscribe, resubscribe | One live cookie per subscription; callbacks occur only while subscribed; successful cookies balance server-held sink references. |
| C09 | Duplicate unsubscribe and controlled `Unadvise` failure | No duplicate sink release; retry/unresolved policy matches Core diagnostics. |
| C10 | Finalize the last subscribed source | Finalizer schedules work only; CCW stays alive until owner-STA `Unadvise` finishes. |
| C11 | Dispatcher stops before pending event teardown | No arbitrary-thread COM cleanup; one unresolved-teardown diagnostic identifies identity, cookie, and apartment. |
| C12 | Native driver retains and later releases a C# callback CCW | Calls succeed only during the documented native hold; native sink/callback reference is balanced independently of RCWs. |
| C13 | Native worker invokes callback across apartments | Delivery and reentry obey COM marshaling and configured dispatcher policy without deadlock. |
| C14 | Nested callbacks pass the same and distinct objects | Callback depth and identity coalescing are correct at every level; temporary callback scopes unwind in reverse order. |

### D. Coexistence with C# COM interop views

| ID | Scenario | Required oracle |
| --- | --- | --- |
| D01 | Hold `[ComImport] IRoot`, wrap its RCW in NetOffice, dispose NetOffice first | Early-bound alias remains callable and returns the same object ID. |
| D02 | Create NetOffice first, obtain an early-bound interface, dispose NetOffice | Early-bound alias remains callable; raw escape is recorded when applicable. |
| D03 | Hold `dynamic` and `object` aliases from `UnderlyingObject` | Escape diagnostic precedes return; both aliases remain callable after NetOffice disposal. |
| D04 | Separate activations return the fixture singleton to NetOffice and early-bound callers | Canonical identity agrees; neither view disconnects the other. |
| D05 | Early-bound alias and NetOffice wrapper receive different interfaces of one tear-off family | Canonical identity agrees despite pointer inequality; calls through both succeed. |
| D06 | C# CCW callback and NetOffice event-sink CCW are live simultaneously | Each native hold is independently advised/retained/released; RCW cleanup does not affect CCW ownership. |
| D07 | External code forcibly calls `ReleaseComObject` in an isolated negative row | Result is classified as unsupported external interference; NetOffice must not claim to prevent disconnection. This row never runs in normal cleanup. |

### E. Object graphs, collections, and return shapes

| ID | Scenario | Required oracle |
| --- | --- | --- |
| E01 | Dispose parent while retaining child and grandchild | Retained descendants remain callable; parent provenance is not ownership. |
| E02 | Dispose child before parent | Parent remains callable; child rejects new calls. |
| E03 | Traverse parent-child cycle repeatedly | Identity entries remain bounded by unique canonical identities, not path count. |
| E04 | Enumerate repeated `VT_UNKNOWN`/`VT_DISPATCH` items | Repeated items coalesce; disposing enumerator does not invalidate retained items. |
| E05 | Break enumeration early and abandon enumerator | Enumerator becomes collectible; retained collection/items remain callable. |
| E06 | Enumerator fails from `Next` | Activity unwinds and previously retained items remain valid. |
| E07 | Clone, replacement, scope `Track`, and `Detach` | Clone acquires an independent lease; replacement is a move; only tracked-not-detached wrappers are invalidated. |
| E08 | Mixed `SAFEARRAY` with null, scalar, repeated, and distinct objects | Each COM element has correct identity and lifetime; null/scalars create no leases. |
| E09 | Method returns null or unsupported interface | No phantom identity or lease remains; documented managed result/error occurs. |

### F. Calls, concurrency, apartments, and shutdown

| ID | Scenario | Required oracle |
| --- | --- | --- |
| F01 | Dispose a wrapper while its `LongCall` is blocked | Call completes after release barrier; registry drop or destruction cannot precede `CallExited`; later calls are rejected. |
| F02 | Wrapper A calls while wrapper B of the same identity is disposed | A completes and remains callable; identity removal waits for the active call. |
| F03 | Callback and disposal run under both legal deterministic schedules | In-flight callback completes; new work follows shutdown policy; teardown occurs once after callback exit. |
| F04 | Call STA-owned identity from MTA and second STA | Configured mode executes native entry on owner STA; rejection mode fails before native entry. |
| F05 | Shutdown requested reentrantly on owner STA | No self-wait; accepted work drains according to policy. |
| F06 | Normal `OfficeSession` shutdown with forgotten wrappers | Event teardown drains before dispatcher stop; remaining live leases and unresolved connections are accurately reported. |
| F07 | Host exits or is killed with live objects | Validates controller isolation only; it is never reported as reference-balance success. |
| F08 | Repeat attach/call/dispose/callback schedules under seeded stress | No crash, hang, underflow, duplicate destruction, lost `Unadvise`, or invariant failure; seed is replayable. |

## Scenario parameter axes

Do not execute the full Cartesian product. Run every P0 scenario in its explicitly named rows, then use deterministic pairwise generation for compatible secondary axes:

- host architecture: x86, x64;
- target framework: .NET Framework 4.6.2, 4.8;
- build: Release required, Debug diagnostic only;
- activation: CLSID, ProgID where applicable;
- caller: owner STA, second STA, MTA;
- ingress: activation, method, property, indexer, enumerator, callback, external RCW;
- interface shape: `IUnknown`, `IDispatch`, dual interface, tear-off, aggregate;
- completion: success, HRESULT failure, managed callback exception, cancellation by shutdown;
- release path: explicit disposal, scope disposal, finalization, event teardown, process isolation kill; and
- alias: none, independent NetOffice wrapper, early-bound C# interface, `dynamic`, escaped raw object.

Critical race schedules are exhaustive over their small defined schedule set and are not delegated to pairwise selection.

## Baselines and harness qualification

### Baseline cohorts

Each native object family first passes a direct C++ fixture self-test. Managed integration then runs four distinct cohorts:

1. early-bound `[ComImport]` only;
2. raw `object`/`dynamic` RCW only;
3. generated NetOffice only; and
4. mixed NetOffice plus early-bound/raw aliases.

Baselines establish that activation, IDL marshaling, connection points, barriers, and telemetry work. They do not define a golden CLR reference-count sequence for NetOffice.

### Oracle mutation tests

Before the harness gates Core, it must demonstrate that its oracles reject controlled faulty builds or fixture modes:

- key identity by `IAlpha`/`IBeta` pointer instead of canonical `IUnknown`;
- root wrappers strongly in diagnostics;
- recursively invalidate a retained child;
- mark a lease released twice;
- drop an identity while a call is blocked;
- dispose a borrowed callback argument;
- run `Unadvise` during callback delivery;
- leak a connection cookie;
- execute `Unadvise` on the finalizer thread; and
- include a forbidden `ReleaseComObject` call in a qualification assembly.

A mutation that survives means the corresponding oracle is not trustworthy and blocks adoption of the harness.

## Failure classification and artifacts

The controller classifies failures as:

- managed assertion or contract failure;
- unexpected HRESULT or managed exception;
- native invariant failure;
- access violation or other native crash;
- CLR fail-fast;
- startup timeout;
- progress-barrier timeout;
- deadlock or total timeout;
- apartment/thread violation;
- trace overflow/corruption;
- unresolved teardown mismatch; or
- infrastructure/activation failure.

For every failure retain:

- scenario ID, all parameter values, seed, and barrier schedule;
- child command line, architecture, CLR and OS versions, and native module hashes;
- complete native and managed traces;
- last completed barrier;
- stdout, stderr, result manifest, and exit code;
- minidump for crash or hang; and
- controller-side oracle explanation with the first violated sequence relation.

Infrastructure failures are reported separately and do not count as product passes or product failures.

## CI execution plan

### Default test workflow

Every test that loads the native fixture or launches `LifetimeTest.ScenarioHost` must use NUnit category `IntegrationTests`. The existing `.github/workflows/tests.yml` command:

```text
dotnet test --filter "TestCategory != IntegrationTests" Source\NetOffice.sln
```

remains the default `windows-2025` gate. It builds and runs without administrator rights, COM registration, or native integration activation. Pure managed controller/oracle tests and the static forbidden-call analyzer should remain outside `IntegrationTests` when they do not load native code.

### Opt-in native integration gate

A dedicated workflow or explicit local command selects `TestCategory=IntegrationTests`. Its minimum row is:

- x64 fixture and x64 host with exact process/DLL bitness matching;
- Release, target .NET Framework 4.6.2, on the pinned Windows image;
- all P0 deterministic native scenarios;
- fixture/oracle qualification tests; and
- no real Office dependency.

The opt-in workflow runs on `workflow_dispatch` and the scheduled gate. It must not require elevation or mutate either machine-wide or per-user COM registration.

### Scheduled gate

- x86 fixture with x86 host and x64 fixture with x64 host;
- every P0 and P1 scenario;
- pairwise secondary axes;
- seeded stress after deterministic races; and
- dump and artifact upload on failure.

.NET Framework 4.x is an in-place runtime family. Testing execution on an actual 4.6.2 CLR versus 4.8 requires separate pinned Windows images; changing only `TargetFrameworkVersion` is not a separate runtime test. Both target-framework builds remain useful for API compatibility.

### Office compatibility lane

Run Excel, Word, Outlook, PowerPoint, and Access only where the matching Office bitness is installed. Cover event-heavy and enumeration-heavy object graphs, mixed PIA/NetOffice aliases, explicit `Quit`, and session shutdown. This lane validates compatibility, not precise reference accounting, and must not use Office process disappearance as a lifetime oracle.

## Advisor review disposition

An independent architecture advisor reviewed the production lifetime contract and proposed harness. The final design incorporates the advisor's merge-blocking requirements:

1. one isolated, architecture-matched child process per native scenario;
2. out-of-band native telemetry and offline oracles;
3. consumer-visible alias liveness instead of RCW internal-count assertions;
4. deterministic barriers instead of sleeps;
5. actual generated wrappers and sinks instead of only direct `COMObject` tests;
6. an independent IL scan for forbidden manual RCW-release APIs;
7. canonical-identity, aggregation, collection, callback, race, and apartment fixtures;
8. disciplined weak-reference and finalizer tests;
9. Job Object deadlines, dump capture, and replay artifacts; and
10. x86/x64, Release JIT, STA/MTA, and separately pinned .NET Framework runtime rows.

Review disposition: acceptable only when all P0 scenarios and harness qualification mutations pass.

## Acceptance criteria

The test harness architecture is implemented when:

- x86 and x64 native fixtures activate without machine-wide registration;
- every native-dependent test is categorized `IntegrationTests`, remains excluded by the default `tests.yml` filter, and runs only through an explicit native integration selection;
- each integration scenario runs in a fresh Job Object child and cannot poison another row;
- the generated fixture API exercises normal NetOffice invocation, enumeration, and event code paths;
- native telemetry observes object identity, lifetime, connection, callback, thread, and barrier facts without COM calls;
- managed diagnostics correlate wrappers, identities, leases, ingresses, calls, callbacks, connections, escapes, and teardown without rooting objects;
- the IL scan rejects any Core or generated call to `ReleaseComObject` or `FinalReleaseComObject`;
- all A-D required user scenarios and expanded P0 scenarios pass with deterministic oracles;
- every explicit identity probe is balanced in `finally` and every fixture object avoids underflow and duplicate destruction;
- live NetOffice, early-bound, raw, and callback aliases remain callable when another independent wrapper is disposed;
- retained callback objects remain callable after callback return;
- event sink holds balance by cookie and `Unadvise` never overlaps callback delivery;
- no identity removal or native destruction precedes an active call or callback exit;
- apartment-sensitive operations enter native code on the required thread or reject before entry;
- GC tests follow the no-inlining, weak-reference, keep-alive, pumped-quiescence protocol;
- the oracle mutation suite kills every required faulty mode; and
- every failure produces enough stable evidence to distinguish product, native fixture, activation, timeout, and infrastructure defects.

## Normative references

- [NetOffice Core COM Lifetime Architecture](Core-COM-Lifetime-Architecture.md)
- [Microsoft: Implementing Reference Counting](https://learn.microsoft.com/en-us/windows/win32/com/implementing-reference-counting)
- [Microsoft: Rules for Implementing QueryInterface](https://learn.microsoft.com/en-us/windows/win32/com/rules-for-implementing-queryinterface)
- [Microsoft: Runtime Callable Wrapper](https://learn.microsoft.com/en-us/dotnet/standard/native-interop/runtime-callable-wrapper)
- [Microsoft: Marshal.ReleaseComObject](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.interopservices.marshal.releasecomobject?view=netframework-4.8.1)
- [Microsoft: COM Apartments](https://learn.microsoft.com/en-us/windows/win32/com/processes--threads--and-apartments)
- [Microsoft: IConnectionPoint::Advise](https://learn.microsoft.com/en-us/windows/win32/api/ocidl/nf-ocidl-iconnectionpoint-advise)
