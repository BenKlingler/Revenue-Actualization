# Revenue-Actualization
## Overview

This suite of Excel VBA macros automates various financial management tasks, including data transfer, forecast extensions, and financial summaries within a workbook that contains "UPCOMING PROJECTS" and "Revenue Summary" sheets.

## Macro Execution Sequence

### Macro: `ExecuteInSequence`

Runs several macros in a specific order to update the workbook's financial data comprehensively.

- `RowInsertandForecastExtension`
- `ActualizeRevenueSummary`
- `RunAllMacros` (which in turn runs `ExtendForecastValues`, `UpdateTotal2025`, and `UpdateRemainingContractAmount`)

### Function: `IsMonthReadyToBeActualized`

Checks if the current month is ready to be actualized based on the dates present in the "Revenue Summary" sheet.

### Macro: `RowInsertandForecastExtension`

Inserts a new column for forecast extension if the current month has passed and the "Revenue Summary" sheet is ready for actualization.

### Macro: `ActualizeRevenueSummary`

Updates the "Revenue Summary" sheet by actualizing previous month's forecast data and adjusting the remaining contract amount accordingly.

### Macro: `ExtendForecastValues`

Extends the forecast values to the new month column in the "Revenue Summary" sheet.

### Macro: `UpdateTotal2025`

Calculates and updates the total values for the year 2025 in the "Revenue Summary" sheet.

### Macro: `UpdateRemainingContractAmount`

Adjusts the remaining contract amount by accounting for the actualized values and distributes any adjustments over the remaining forecast months.

## Helper Functions

- `FindNextVisibleRow`: Finds the next visible row within a specified column, ensuring that the row is not hidden and contains data.
