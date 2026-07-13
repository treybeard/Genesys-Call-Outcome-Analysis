# LRS Spanish Outcome Logic

LRS Spanish exports have a different outcome classification logic than Switchboard and Non-Atty flows.

## Outcome Classification

| Category | Condition |
|----------|-----------|
| **Successful** | `Outcome Success == 1` AND `Abandoned != 'YES'` |
| **Failed** | `Outcome Success != 1` AND `Abandoned != 'YES'` |
| **Abandoned** | `Abandoned == 'YES'` (regardless of Outcome Success) |

## Rationale

In LRS Spanish exports, a call can be both "successful" (correct routing/path completed) AND "abandoned" (user hung up before completion). The system tracks both dimensions independently.

Therefore:
- Successful calls are those that completed successfully AND the user did NOT abandon
- Failed calls are those where Outcome Success != 1 AND user did NOT abandon
- Abandoned calls include ALL calls where `Abandoned == 'YES'`, regardless of outcome value

## Verification Example

| Metric | Formula | Example |
|--------|---------|---------|
| Total Calls | `len(df)` | 1,043 |
| Successful | `((df['Outcome Success'] == 1) & (df['Abandoned'] != 'YES')).sum()` | 878 |
| Failed | `((df['Outcome Success'] != 1) & (df['Abandoned'] != 'YES')).sum()` | 17 |
| Abandoned | `(df['Abandoned'] == 'YES').sum()` | 148 |
| Sum Check | `successful + failed + abandoned == total` | ✓ |

## Pitfall: Simple Subtraction Fails

The formula `Failed = Total - Successful - Abandoned` does NOT work for LRS Spanish because some rows are counted in both `Successful` and `Abandoned` (users who completed their call but still hung up before final wrap-up).

Always use the conditional logic above for LRS Spanish exports.
