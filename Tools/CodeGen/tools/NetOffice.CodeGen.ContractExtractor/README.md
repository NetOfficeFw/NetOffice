# NetOffice wrapper contract extractor

Copyright (c) 2026 NetOffice contributors. SPDX-License-Identifier: MIT.

This small, dependency-free .NET CLI extracts a deterministic source contract from
`Source/<Api>`. It deliberately records what can be established from source text
and records unknowns rather than guessing. The output is intended as an input
contract for the independent code-generation pipeline, not as generated wrapper
source.

## Invocation

From the `NetOffice` repository root:

```text
dotnet run --project Tools/CodeGen/tools/NetOffice.CodeGen.ContractExtractor/NetOffice.CodeGen.ContractExtractor.csproj -- \
  --source Source --api DAO \
  --output Tools/CodeGen/contracts/wrapper/DAO.wrapper-contract.json
```

The command writes the contract and, beside it, `<Api>.compatibility-ledger.json`
and `<Api>.classification.json`. Use `--ledger` and `--classification` to choose
explicit paths. `--help` prints the complete option synopsis.

## Contract shape

`contracts/wrapper/wrapper-contract.schema.json` defines schema version `1.0`.
A contract contains the normalized source roots and sorted source-file hashes,
type records, public/protected member signatures, attributes, base types,
interfaces, safely recognizable invocation text, and explicit `Unknowns` and
`Ambiguities` arrays. File and record ordering is ordinal and output is UTF-8
without a BOM with a final newline. Source paths use `/` separators, and source
hashes are calculated over normalized LF text.

The checked-in DAO files are a reproducible example and can be regenerated with
the command above. The extractor never reads network resources, compiled
assemblies, or product-specific configuration.
