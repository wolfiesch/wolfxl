# Versioning and Upgrades

## Upgrade policy

- Pin WolfXL version in production environments.
- Review changelog and known limitations before upgrading.
- Re-run core workbook regression tests on each upgrade.

## Recommended upgrade flow

1. Upgrade in staging
2. Run fidelity/performance smoke suite
3. Validate output in Excel for business-critical templates
4. Roll out incrementally

## Compatibility note

- Primary native module: `wolfxl._rust`
- Legacy compatibility module: `excelbench_rust` shim
