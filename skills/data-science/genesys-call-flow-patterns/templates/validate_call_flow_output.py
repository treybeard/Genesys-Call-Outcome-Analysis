# Genesys Call Flow Validation Script

This script validates the output of call flow analysis for Switchboard and Non-Atty flows.

## Usage
```bash
python3 validate_call_flow_output.py /path/to/output.xlsx
```

## What It Checks
1. Total calls in Summary matches row count
2. Success + failed + abandon ≈ total (within 0.1% tolerance)
3. Avg Duration is realistic (< 24 hours for average)
4. Call Paths sheet has at least one entry
5. No NaN values in critical columns

## Expected Output Format
```
✅ Total calls: 3,257
✅ Success + failed + abandon = 100.0%
✅ Avg Duration: 6:47.89 (realistic)
✅ Call Paths entries: 65
```