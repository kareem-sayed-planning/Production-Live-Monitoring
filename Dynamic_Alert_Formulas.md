# ⚡ Dynamic Alert Generation Formulas (Excel LET & LAMBDA)

To ensure the dashboard UI remains lightning-fast and isn't bogged down by constant VBA macro executions, the visual warning banners are powered purely by advanced Excel Dynamic Arrays (`LET`, `MAP`, `LAMBDA`). 

This logic automatically translates raw variance metrics into readable operational warnings.

## 1. Global System Status Engine
*Evaluates the maximum loss across all production stages and dictates the global dashboard header.*

```excel
=LET(
    Vars, ROUND($AX$68:$BA$68, 0),
    Losses, IF(Vars < 0, ABS(Vars), 0),
    MaxLoss, MAX(Losses),

    IFS(
        MaxLoss = 0, "✅ ALL SHOPS NORMAL",
        MaxLoss > 3, "🚨 CRITICAL ISSUES",
        TRUE, "⚠️ MINOR ISSUES"
    )
)
