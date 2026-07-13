# ATTY Outcome Logic

ATTY (Attorney) exports follow the same outcome classification as Switchboard and Non-Atty flows.

## Outcome Classification

| Category | Condition |
|----------|-----------|
| **Successful** | `Outcome Success == 1` AND `Abandoned != 'YES'` |
| **Failed** | `Outcome Success != 1` AND `Abandoned != 'YES'` |
| **Abandoned** | `Abandoned == 'YES'` (regardless of Outcome Success) |

## Rationale

In ATTY exports, the system tracks both outcome completion AND abandonment independently. A call can complete successfully but still be marked as abandoned if the user hung up before final wrap-up.

Therefore:
- Successful calls are those that completed successfully AND the user did NOT abandon
- Failed calls are those where Outcome Success != 1 AND user did NOT abandon
- Abandoned calls include ALL calls where `Abandoned == 'YES'`, regardless of outcome value

## Verification Example

| Metric | Formula | Example (April) |
|--------|---------|-----------------|
| Total Calls | `len(df)` | 1,164 |
| Successful | `((df['Outcome Success'] == 1) & (df['Abandoned'] != 'YES')).sum()` | 971 |
| Failed | `((df['Outcome Success'] != 1) & (df['Abandoned'] != 'YES')).sum()` | 153 |
| Abandoned | `(df['Abandoned'] == 'YES').sum()` | 40 |
| Sum Check | `successful + failed + abandoned == total` | ✓ |

## Pitfall: Simple Subtraction Fails

The formula `Failed = Total - Successful - Abandoned` does NOT work because some rows are counted in both `Successful` and `Abandoned`.

Always use the conditional logic above for ATTY exports.