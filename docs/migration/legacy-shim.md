# Legacy `excelbench_rust` Shim

WolfXL now uses `wolfxl._rust` as the primary native module name.

To preserve compatibility for existing environments, a shim package named `excelbench-rust` re-exports `wolfxl._rust`.

## Guidance

- New code: import/use WolfXL interfaces (`from wolfxl import ...`).
- Existing legacy code: can continue to work via shim during migration.

## Migration target

- Primary runtime module: `wolfxl._rust`
- Legacy compatibility module: `excelbench_rust` (shim)

## Validation command

```bash
python3 -c "import wolfxl._rust as m; print(m.build_info())"
```

If you still depend on legacy imports, keep shim installed until the migration is complete.
