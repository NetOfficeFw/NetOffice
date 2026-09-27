# NetOffice PowerPoint COM Performance

## Summary

Slido uses NetOffice 1.9.9 for PowerPoint COM integration. NetOffice adds
measurable overhead to individual COM calls, but the largest observed slowdown
comes from how NetOffice collections are consumed in Slido.

Two loops call LINQ `Count()` on NetOffice COM collections:

- `PresentationEditor.GetSlideContent()` calls `shapes.Count()`.
- `PresentationEditor.GetSlidesContent()` calls `slides.Count()`.

NetOffice collections implement `IEnumerable<T>`, but not `ICollection<T>`.
LINQ therefore calculates `Count()` by enumerating the complete COM collection.
Calling it in a loop condition repeats that enumeration for every item, creates
thousands of temporary COM wrappers, and changes a linear traversal into an
$O(n^2)$ traversal.

The first optimization should be reducing COM call and proxy counts. Replacing
NetOffice with direct PowerPoint interop is not justified before these issues
are addressed.

## Affected Code

The known quadratic loops are in
`src/SlidoCore/PowerPoint/Editor/PresentationEditor.cs`:

- `GetSlideContent()` at the `shapes.Count()` loop condition.
- Private `GetSlidesContent()` at the `slides.Count()` loop condition.

The presentation-content path is exposed through
`SidebarHostApi.GetPresentationAiModelData()` and runs synchronously on the
main dispatcher through `DispatcherHelper.InvokeOnMainThreadAsync()`. Redundant
COM work on this path directly extends the time for which PowerPoint's UI thread
is occupied.

## Measured Impact

A temporary benchmark exercised a real PowerPoint 16.0 process through direct
PowerPoint PIA interop and NetOffice 1.9.9. It created a blank presentation with
80 rectangle shapes. PowerPoint startup was excluded. Reported durations are
medians over three to five rounds.

| Operation | Direct PIA | NetOffice | Observation |
| --- | ---: | ---: | --- |
| 5,000 cached `Shapes.Count` reads | 586.5 ms | 1,845.4 ms | NetOffice was 3.15 times slower |
| 5,000 cached `Shape.Name` reads | 620.8 ms | 1,866.6 ms | NetOffice was 3.01 times slower |
| Optimized 80-shape traversal | 175.8 ms | 111.6 ms | Object cleanup strategies make this comparison less reliable |
| Traversal using `shapes.Count()` | Not measured | 14,069.7 ms | 126 times slower than optimized NetOffice traversal |
| Proxies added by one traversal | Not applicable | 6,642 versus 81 | 82 times more proxies |
| Peak open NetOffice proxies | Not applicable | 6,649 versus 9 | Large temporary proxy tree |

The cached scalar property reads are the cleanest direct-interop comparison.
They show approximately three times the per-call cost for NetOffice in this
environment. Object creation and traversal comparisons depend on different PIA
and NetOffice cleanup mechanisms and should not be interpreted as a general
claim that either library is faster.

The benchmark is sufficient to isolate the quadratic traversal problem. It is
not a release-quality comparison across Office versions, machines, or document
types.

## Why `Count()` Is Expensive

NetOffice's `Shapes.Count` property performs one COM property read. By contrast,
LINQ `shapes.Count()` uses the collection's custom enumerator. NetOffice obtains
PowerPoint's `_NewEnum`, calls `MoveNext` and `Current`, and creates a managed
NetOffice wrapper for every returned shape.

For 80 shapes, the current loop produces exactly 6,642 proxies:

1. The loop condition is evaluated 81 times.
2. Each `Count()` call creates one enumerator and 80 shape wrappers.
3. The loop body creates another 80 wrappers through `shapes[i]`.
4. Accessing `slide.Shapes` creates one collection wrapper.

The resulting count is:

$$
81 \times (1 + 80) + 80 + 1 = 6{,}642
$$

The optimized loop creates one `Shapes` wrapper and 80 indexed `Shape`
wrappers, for a total of 81. Disposing each shape at the end of its iteration
also keeps the peak number of open proxies low.

Relevant NetOffice 1.9.9 implementation details:

- [`Shapes` implements `IEnumerableProvider<Shape>` and exposes a separate `Count` property](https://github.com/NetOfficeFw/NetOffice/blob/da3f0f1c22a47c0787325cca98e4d1eb86d0d8bf/Source/PowerPoint/DispatchInterfaces/Shapes.cs)
- [`IEnumerableProvider<T>` extends only `IEnumerable<T>`](https://github.com/NetOfficeFw/NetOffice/blob/da3f0f1c22a47c0787325cca98e4d1eb86d0d8bf/Source/NetOffice/CollectionsGeneric/IEnumerableProvider.cs)
- [The custom enumerator uses `_NewEnum` and creates a wrapper for each item](https://github.com/NetOfficeFw/NetOffice/blob/da3f0f1c22a47c0787325cca98e4d1eb86d0d8bf/Source/NetOffice/Utils.cs)

## NetOffice Per-Call Overhead

NetOffice 1.9.9 dispatches COM properties and methods through reflection-based
`Type.InvokeMember` calls. It also validates wrappers, manages parent-child
proxy relationships, translates exceptions, and optionally records performance
data.

This work explains the measured overhead on cached scalar getters. The cost is
paid at every COM boundary, so minimizing calls matters more than optimizing
ordinary managed code around them.

Relevant implementation:

- [`Invoker.PropertyGet()` delegates to `Type.InvokeMember`](https://github.com/NetOfficeFw/NetOffice/blob/da3f0f1c22a47c0787325cca98e4d1eb86d0d8bf/Source/NetOffice/Invoker.cs)

## Additional High-Frequency COM Patterns

### Repeated Shape Classification

`src/SlidoCore/Services/DocumentApi/Extensions/ShapeExtensions.cs` implements
`IsTable()`, `IsChart()`, and `IsSmartArt()` as separate methods. Calling all
three can read `shape.Type` up to six times for one ordinary shape. Placeholder
classification also repeatedly retrieves `PlaceholderFormat` and
`ContainedType`.

Read `shape.Type` once per shape. Retrieve and dispose `PlaceholderFormat` only
when the cached type is `msoPlaceholder`.

### Repeated Indexer Retrieval

The following code repeatedly retrieves wrappers for the same COM objects:

- `DocumentApi.GetSlidoPresentation()` evaluates `slides[i]` multiple times.
- `PresentationEditor.GetSmartArtContent()` repeatedly evaluates `nodes[i]`,
  `nodes[i].Shapes`, and `nodes[i].TextFrame2`.
- `PresentationEditor.GetPresenterNotes()` creates a `notes` wrapper but then
  repeatedly accesses `slide.NotesPage` instead of using it.

Retrieve each indexed COM object once, keep it within its parent collection's
lifetime, and dispose it at the end of the iteration.

### Chained Property Access

Expressions such as the following cross several COM boundaries and create
intermediate wrappers:

```csharp
cell.Shape.TextFrame2.TextRange.Text
```

Cache and dispose intermediate objects when a loop uses them repeatedly. This
is especially important for tables, SmartArt nodes, text ranges, and
placeholder formats.

### Per-Cell Chart Reads

`PresentationEditor.GetChartContent()` calls
`worksheet.Cells[row, column].Value` for every chart cell. Each iteration
retrieves COM objects independently.

Prefer reading a rectangular range's `Value2` into an `object[,]` in one COM
call, then process the values in managed memory. This reduces both NetOffice
and PowerPoint overhead.

### Repeated Collection Properties

Loops over tables and notes repeatedly access collection properties such as
`table.Rows`, `table.Columns`, and `slide.NotesPage`. Cache the collection
wrapper and its `Count` before entering the loop.

## Recommended Loop Pattern

Use the COM `Count` property once and dispose each child wrapper within the
parent collection's scope:

```csharp
using var shapes = slide.Shapes;
var shapeCount = shapes.Count;

for (var i = 1; i <= shapeCount; i++)
{
    using var shape = shapes[i];

    // Read the shape here.
}
```

Apply the same pattern to slides, SmartArt nodes, table rows and columns, and
other NetOffice collections.

## Lifetime Management

NetOffice tracks COM wrappers in parent-child relationships. Disposing a parent
recursively disposes its children. A child obtained from a collection must not
outlive that collection:

```csharp
using var slides = presentation.Slides;
using var slide = slides[1];

// The slide is valid within this scope.
```

Do not dispose `slides` and continue using `slide`. The benchmark encountered
an `ObjectDisposedException` when a child presentation or slide was retained
after its parent collection had been disposed.

Explicitly disposing each loop item is still useful even though disposing the
collection eventually disposes all descendants. Per-item disposal removes the
wrapper from the parent proxy tree earlier and prevents high transient proxy
counts.

Do not disable NetOffice proxy management as a performance shortcut. That
would exchange measured overhead for fragile COM lifetime handling and
potential leaks.

NetOffice's recursive disposal implementation is in
[`COMObject.Dispose()`](https://github.com/NetOfficeFw/NetOffice/blob/da3f0f1c22a47c0787325cca98e4d1eb86d0d8bf/Source/NetOffice/COMObject.cs).

## Recommended Work Order

1. Replace the two LINQ `Count()` loop conditions with cached COM `Count`
   values.
2. Dispose indexed slide and shape wrappers at the end of each iteration.
3. Cache shape type and placeholder metadata during content extraction.
4. Remove repeated `nodes[i]`, `slides[i]`, and `slide.NotesPage` retrievals.
5. Cache table and notes collections and counts.
6. Batch chart worksheet reads through `Range.Value2`.
7. Benchmark `GetPresentationContent()` on representative customer decks.
8. Compare NetOffice and direct PIA interop again only after both paths perform
   the same logical COM operations and use equivalent lifetime policies.

## Diagnostic Instrumentation

Measure user-visible operations with `Stopwatch`, including:

- `PresentationEditor.GetPresentation()`
- `PresentationEditor.GetPresentationContent()`
- `DocumentApi.GetSlidoPresentation()`
- Per-slide content extraction

Record slide count, total and maximum shapes per slide, total elapsed time,
NetOffice proxy additions, and peak `Core.Default.ProxyCount`.

NetOffice's built-in `PerformanceTrace` can inventory individual calls:

```csharp
var trace = NetOffice.Core.Default.Settings.PerformanceTrace;
var powerPointCalls = trace["NetOffice.PowerPointApi"];

powerPointCalls.IntervalMS = 0;
powerPointCalls.Enabled = true;

trace.Alert += (_, args) =>
{
    // Aggregate by EntityName, MethodName, and CallType.
};

trace.Enabled = true;
```

Tracing every call changes the timings. Use it to identify call counts and hot
members, then disable it when measuring end-to-end duration. Proxy events also
add overhead because NetOffice constructs ownership information for event
subscribers. Collect proxy counts separately from the timing runs.

## Conclusion

NetOffice adds measurable per-call overhead, but the current performance risk
is dominated by excessive COM calls and temporary proxy creation. The
`Count()` loops are the highest-priority issue: an 80-shape traversal measured
14.1 seconds and created 6,642 proxies, while the cached-`Count` version
measured 111.6 milliseconds and created 81 proxies.

Optimize COM access patterns before considering an interop-library migration.
