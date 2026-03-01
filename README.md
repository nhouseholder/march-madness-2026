# 🏀 March Madness 2026 — Predictive Bracket

> Multi-source analytical bracket prediction for the 2026 NCAA Men's Basketball Tournament.  
> **Projected Champion: Duke Blue Devils**

## 📊 Data Sources

| Source | Type | Coverage |
|--------|------|----------|
| **KenPom** | Adjusted Efficiency Ratings | 365 teams, full 2025-26 season |
| **Haslametrics** | Team Ratings + Bracketology | 365 teams + projected seeds |
| **EvanMiya** | BPR, Kill Shot, Relative Ratings, Under-seeded, Efficiency Landscape | Instagram graphs (8 total) |
| **CBB Analytics** | Versatility Index Grade (VIG) | Instagram graph |
| **Bracket Science** | Championship Profile (Off-Def Quadrant) | Instagram graph |
| **Dr. Locks** | Graph to Greatness (Trajectory Milestones) | Instagram graph |
| **NCAA Historical** | Seed pairing results 1985-2025 | 40 years, all rounds |

## 🏆 Final Four

| Semifinal 1 | Semifinal 2 |
|-------------|-------------|
| **(1) Duke** vs (1) Arizona | **(1) Michigan** vs (2) Illinois |

**Championship: Duke 74, Michigan 68**

## 📁 Contents

| File | Description |
|------|-------------|
| `bracket.html` | Interactive visual bracket (dark theme, 4 regions, Final Four, seeding table, upset picks, research sections) |
| `March_Madness_2026_Prediction_Report.docx` | Comprehensive Word report: all data, sources, theories, projections, and 12 suggestions for improvement |
| `generate_report.py` | Python script that generates the Word document |

## 🔍 Key Upsets Projected

- **(9) Clemson** over (8) Miami FL — 9-seeds lead 83-77 all-time
- **(9) Missouri** over (8) UCF — SEC battle-tested
- **(9) NC State** over (8) Iowa — ACC gauntlet prep
- **(10) UCLA** over (7) Villanova — Big Ten SOS gap
- **(3) Gonzaga** over (2) Nebraska — March DNA vs first deep run
- **(2) Illinois** over (1) UConn — #1 offensive efficiency (133.99) in Elite 8

## 🎯 Methodology

Bracket built via weighted synthesis:
1. **KenPom AdjEM** as primary ranking input
2. **Haslametrics bracketology** for seed-line placement cross-reference
3. **Graph-derived qualitative analysis** for championship profiling (BPR, VIG, Kill Shot, Championship Zone, Graph to Greatness)
4. **Historical seed data** for upset calibration (40 years of results)
5. **8 validated predictive theories** (defensive efficiency, FT rate, turnover margin, experience, tempo control, 3PT variance, SOS gap, conference tournament fatigue)

## 🚀 View the Bracket

```bash
cd march-madness-2026
python3 -m http.server 8765
# Open http://localhost:8765/bracket.html
```

---
*Generated March 1, 2026 by GitHub Copilot (Claude Opus 4.6)*
