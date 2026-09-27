using NetOffice.CodeGen.Application;

const string Help = "NetOffice Code Generator\n\nUsage: netoffice-codegen <command> [options]\n\nCommands:\n  import-typelib       Import an explicit Office type library (not locked)\n  merge                Merge observations into the data graph\n  docs sync            Synchronize pinned VBA documentation\n  generate             Generate wrappers; --locked supports --check\n  verify               Verify pinned inputs and generated tree\n  diff                 Compare --source with --actual\n  explain <id>         Explain projection for a logical id\n  bootstrap-ownership  Check or apply initial ownership\n\nOptions: --locked --check --apply --data <dir> --docs <dir> --docs-profile <name>\n         --source <dir> --output <dir> --actual <dir> --report <file> --id <id> --help\nExit codes: 0 success/clean, 1 validation/failure, 2 drift, 130 cancelled.";
if (args.Length == 0 || args.Contains("--help", StringComparer.OrdinalIgnoreCase)) { Console.WriteLine(Help); return 0; }
var command=args[0].ToLowerInvariant(); var start=1;
if(command=="docs" && args.Length>1 && args[1].Equals("sync",StringComparison.OrdinalIgnoreCase)) { command="docs"; start=2; }
string? data=null,docs=null,source=null,output=null,id=null,docsProfile=null,report=null; var locked=false; var check=false; var apply=false;
for(var i=start;i<args.Length;i++) {
    switch(args[i]) {
      case "--locked": locked=true; break; case "--check": check=true; break; case "--apply": apply=true; break;
      case "--data": data=Value(args,ref i); break; case "--docs": docs=Value(args,ref i); break; case "--docs-profile": docsProfile=Value(args,ref i); break; case "--report": report=Value(args,ref i); break; case "--source": source=Value(args,ref i); break;
      case "--output": output=Value(args,ref i); break; case "--actual": output=Value(args,ref i); break; case "--id": id=Value(args,ref i); break;
      default: if(command=="explain"&&id is null&&!args[i].StartsWith('-')) id=args[i]; else { Console.Error.WriteLine($"Unknown option '{args[i]}'."); return 1; } break;
    }
}
using var cts=new CancellationTokenSource(); Console.CancelKeyPress += (_,e)=>{e.Cancel=true;cts.Cancel();};
CommandResult result;
try { result=ApplicationService.Execute(new(command,data,docs,source,output,id,docsProfile,report,locked,check,apply),cts.Token); }
catch (ArgumentException ex) { Console.Error.WriteLine(ex.Message); return 1; }
if(result.ExitCode==0) Console.WriteLine(result.Message); else Console.Error.WriteLine(result.Message);
return result.ExitCode;
static string Value(string[] a,ref int i) { if(++i>=a.Length||a[i].StartsWith('-')) throw new ArgumentException($"Option {a[i-1]} requires a value."); return a[i]; }
