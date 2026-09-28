using NetOffice.CodeGen.Application;

const string Help = """
NetOffice Code Generator

Usage: netoffice-codegen <command> [options]

Commands:
  import-typelib       Import an explicit Office type library (not implemented)
  merge                Merge observations into the data graph (not implemented)
  docs sync            Synchronize pinned VBA documentation (not implemented)
  generate             Project Data v2 and emit wrappers
  verify               Verify pinned inputs and generated tree (not implemented)
  diff                 Compare complete --expected and --actual trees
  explain <id>         Print projection traces for a logical id
  bootstrap-ownership  Check or apply initial ownership (not implemented)

Generation options:
  --data <file|dir>          Canonical Data v2 graph.json
  --contract <file|dir>      Wrapper Contract file or corpus directory
  --policy <file>            Projection policy JSON
  --projects <list|all>      Comma-separated products, or all (default: all)
  --docs <dir>               Offline pinned documentation source
  --docs-profile <name>      baseline or vba
  --output <dir>             Destination root
  --source <dir>             Existing Source root (required for isolated output)
  --isolated-output          Emit a buildable Source layout with classified companions
  --layout <generated|isolated>
  --from-data, --exploratory Permit data-only exploratory projection
  --locked                   Require a contract for every selected product
  --check                    Report drift without mutating output
  --no-cache                 Disable cache (generation is currently always cacheless)
  --report <file>            Deterministic JSON execution report

Diff options:
  --expected <dir> --actual <dir>

Exit codes: 0 success/clean, 1 usage/validation/generation failure,
            2 check or diff drift, 130 cancellation.
""";

if (args.Length == 0 || args.Contains("--help", StringComparer.OrdinalIgnoreCase))
{
    Console.WriteLine(Help);
    return 0;
}

var command = args[0].ToLowerInvariant();
var start = 1;
if (command == "docs" && args.Length > 1 && args[1].Equals("sync", StringComparison.OrdinalIgnoreCase))
{
    start = 2;
}

string? data = null;
string? docs = null;
string? source = null;
string? output = null;
string? id = null;
string? docsProfile = null;
string? report = null;
string? contract = null;
string? policy = null;
string? projects = null;
string? expected = null;
string? actual = null;
var locked = false;
var check = false;
var apply = false;
var noCache = false;
var fromData = false;
var isolatedOutput = false;

try
{
    for (var index = start; index < args.Length; index++)
    {
        switch (args[index])
        {
            case "--locked": locked = true; break;
            case "--check": check = true; break;
            case "--apply": apply = true; break;
            case "--no-cache": noCache = true; break;
            case "--from-data":
            case "--exploratory": fromData = true; break;
            case "--isolated":
            case "--isolated-output": isolatedOutput = true; break;
            case "--data": data = Value(args, ref index); break;
            case "--docs": docs = Value(args, ref index); break;
            case "--docs-profile": docsProfile = Value(args, ref index); break;
            case "--report": report = Value(args, ref index); break;
            case "--source": source = Value(args, ref index); break;
            case "--output": output = Value(args, ref index); break;
            case "--actual": actual = Value(args, ref index); break;
            case "--expected": expected = Value(args, ref index); break;
            case "--contract": contract = Value(args, ref index); break;
            case "--policy": policy = Value(args, ref index); break;
            case "--projects": projects = Value(args, ref index); break;
            case "--id": id = Value(args, ref index); break;
            case "--layout":
            {
                var layout = Value(args, ref index);
                isolatedOutput = layout.Equals("isolated", StringComparison.OrdinalIgnoreCase)
                    ? true
                    : layout.Equals("generated", StringComparison.OrdinalIgnoreCase)
                        ? false
                        : throw new ArgumentException("--layout must be generated or isolated.");
                break;
            }
            default:
                if (command == "explain" && id is null && !args[index].StartsWith("-", StringComparison.Ordinal)) id = args[index];
                else throw new ArgumentException($"Unknown option '{args[index]}'.");
                break;
        }
    }
}
catch (ArgumentException error)
{
    Console.Error.WriteLine(error.Message);
    return 1;
}

using var cancellation = new CancellationTokenSource();
Console.CancelKeyPress += (_, eventArgs) =>
{
    eventArgs.Cancel = true;
    cancellation.Cancel();
};

var result = ApplicationService.Execute(
    new CommandRequest(
        command,
        data,
        docs,
        source,
        output,
        id,
        docsProfile,
        report,
        locked,
        check,
        apply,
        contract,
        policy,
        projects,
        noCache,
        expected,
        actual,
        fromData,
        isolatedOutput),
    cancellation.Token);
if (result.ExitCode == 0) Console.WriteLine(result.Message);
else Console.Error.WriteLine(result.Message);
return result.ExitCode;

static string Value(string[] arguments, ref int index)
{
    if (++index >= arguments.Length || arguments[index].StartsWith("-", StringComparison.Ordinal))
        throw new ArgumentException($"Option {arguments[index - 1]} requires a value.");
    return arguments[index];
}
