# NetOffice.CodeGen.Cli

The executable is intentionally deterministic and offline in locked mode. Use `--help` for commands and options. `generate --locked --check` and `bootstrap-ownership --locked --check` never write files; use an output directory for dry-run generation. Locked commands require explicit existing pinned input directories and do not access Office or the network.

Exit codes are 0 for success, 1 for validation or execution failure, 2 for detected drift, and 130 for Ctrl+C cancellation.
