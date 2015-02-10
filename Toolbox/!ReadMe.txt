[Depricated] -- fixed from artnib in #112250 !!! thx

Dear crew member,

No idea why but sometimes its failed to compile the toolbox.

Error message is:
'Missing compiler required member 'System.Runtime.CompilerServices.ExtensionAttribute..ctor'

I guess we use Linq to fetch over a .NET2 assembly and this cause problems.

hotfix: remove Mono.Cecil from references and add again from 'Toolbox\Libs' folder.
Now it works fine! (its a kind of magic because it works fine for all the time, but NO Toolbox is started in .NET 2 without Linq...)

related topic:
http://stackoverflow.com/questions/4353335/mono-cecil-missing-compiler-required-member-system-runtime-compilerservices-ex

It looks like we need a .NET4 compiled Mono.Cecil assembly but i can't find this one.
If you can help, please let me know.

*Sebastian
[public.sebastian@web.de]

[/Depricated]
