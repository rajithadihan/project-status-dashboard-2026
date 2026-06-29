IPD Data Enrichment Dashboard V4

How to use:
1. Keep Project-status.xlsx in this same folder.
2. Update these sheets in Excel:
   - All: full product scope.
   - Task: completed/updated products.
   - Working Days: dates and engineer worked/leave tracking.
3. Open start_dashboard.bat on Windows or start_dashboard.command on Mac.
4. Use Refresh Data after saving Excel.

Key logic:
- Total products are counted from All sheet.
- Completed products are unique Task product codes that match All sheet and have a completed-like status or a completion date.
- Daily graphs use only dates in the Working Days sheet.
- If an engineer is marked Worked on a date but has no completed Task rows for that date, the graph shows 0 for that engineer.
- Actual team rate = completed products / team working days from Working Days sheet.
- Individual rate = engineer completed products / engineer worked days.
- Product Analytics hierarchy is Category -> Class -> Stock Group.


V5 update:
- Removed Capacity Planner page.
- Added Forecast page for the 30 Sep 2026 deadline.
- Forecast uses only dates available in the Working Days sheet.
- Actual rate is based on elapsed worked days; future dates in Working Days are used for deadline forecasting.
- Forecast page shows Safe Line / Watch Line / At Risk status, required daily rate, current rate, engineer requirement, and suggested action.

V5.1 update:
- Forecast page now calculates remaining working days to 30 Sep 2026 if future work dates are not listed.
- Actual productivity still uses only elapsed Working Days sheet dates.

V6 update:
- Forecast page compressed to a single executive view.
- Key forecast numbers are kept at the top.
- Forecast to Deadline and Daily Output vs Required Line charts are smaller and more compact.
- Day-to-day x-axis spacing reduced on forecast charts.


V7 update:
- Forecast page is compressed to fit one screen without vertical scrolling.
- Removed the forecast extra note strip from the main view.
- Forecast charts are smaller and dense so the page remains presentation-friendly.
- Forecast chart x-axis still shows working dates, but spacing is reduced to avoid a long page.


V8 update:
- Active page navigation button is now highlighted.
- Refresh Data button is now icon-only.


V9 update:
- Full project package rebuilt.
- Active navigation page is clearly highlighted with cyan border, glow, and underline.
- Refresh Data button remains icon-only.


V10 update:
- Added Forecast Planning page.
- Allows future planning periods with from/to dates, extra resources, output per extra resource/day, and manual daily target.
- Shows achievable vs plan gap, projected completion, remaining at deadline, and scenario charts.
- Planning assumptions are saved in browser localStorage and do not change actual dashboard data.


V11 update:
- Removed separate Forecast tab/page.
- Summary page now includes executive combined view with forecast status, recommendation, forecast charts, daily output vs required line, ABC status and key takeaways.
- Engineer Performance, Product Analytics, Recent Tasks and Forecast Planning pages remain unchanged.


V12 update:
- Included the fixed safer lineChart function.
- Reduced date spacing while keeping every working day visible.
- Increased ABC Class Progress table font size.
- Removed active page glow and underline.
- Forecast page remains removed; forecast data is included in Summary.


V13 update:
- Dashboard CSS has been separated from dashboard.html.
- New file added: styles.css
- dashboard.html now references styles.css using:
  <link rel="stylesheet" href="styles.css">

V14 update:
- Forecast Planning page redesigned and starts blank.
- Removed Manual Daily Target.
- Planning inputs: From Date, To Date, Per Day Estimation, Remove.
- One planning line represents one engineer/resource.
- Multiple lines can be added; overlapping dates are summed.
- Summary page shows Balance Products on Deadline based on current rate.
- Recommendation box removed from Summary.


V15 update:
- Logo box updated to white background, 65px x 65px.
- Removed the full Forecast Outcome bar from Summary.
- Balance Products on Deadline now replaces the Recommendation box.
- Forecast Planning now uses current actual rate as the baseline.
- Added engineer lines are treated as additional output on top of the current plan.
