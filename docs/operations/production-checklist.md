# Production Rollout Checklist

## Before rollout

1. Validate compatibility on representative workbooks
2. Compare output files with current production path
3. Benchmark runtime and memory under realistic load
4. Confirm fallback path if migration issue appears

## During rollout

1. Start with one low-risk workflow
2. Monitor runtime, failure rate, and file-integrity signals
3. Keep side-by-side verification for initial deployments

## After rollout

1. Track regressions by workbook type
2. Expand gradually to additional pipelines
3. Capture lessons in release notes/changelog
