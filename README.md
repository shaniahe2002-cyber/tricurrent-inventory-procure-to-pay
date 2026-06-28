# Tricurrent Inventory and Procure-to-Pay Automation Suite

Excel VBA system for a fictional coffee roasting company, covering inventory management, purchase requisitioning, finance approval, vendor PO generation, and AP reconciliation.

## Problem

A growing coffee roasting operation across three companies sharing one warehouse needed a way to track 90+ SKUs, automatically generate purchase requisitions when stock ran low, route them through finance approval, and reconcile what inventory received against what AP had on record, all without a paid ERP system.

## Tools

Excel VBA, Power Query, multi-step UserForms, FIFO lot costing logic

## Key Features

- Inventory Management system with FIFO costing, cycle counts, and bundle-based production tracking
- Purchase Requisition automation that detects reorder points and auto-generates requisitions
- Finance Approval Form with Approve/Deny workflow and audit logging
- Automatic vendor PO generation, one PO sheet per unique vendor
- Bundle Usage Expected vs Actual variance detection, comparing production batch counts against real dispense transactions, with returns netted out
- Live Stock reconciliation check, comparing cumulative Inflow minus Outflow against the live stock balance to catch discrepancies
- AP Log with Power Query-based reconciliation against an Inventory-exported staging file

## What's in This Repo

| File | Description |
|---|---|
| `TricurrentInventoryGroup_Inventory_Management.xlsm` | Main inventory system: Inflow, Outflow, Live Stock, Bundle List, Batch Bundle Usage, dashboard |
| `Purchase_Requisition.xlsm` | Auto-generates purchase requisitions from low-stock inventory, routes through Finance Approval, generates vendor POs |
| `AP_Log.xlsx` | Accounts Payable reconciliation log, pulls from Inflow_Staging.xlsx via Power Query |
| `Inflow_Staging.xlsx` | Handoff file exported by the Inventory system for AP to reconcile against |
| `TriCurrent Inventory Group Logo.png` | Logo used in the Finance Approval Form header |

## Setup Notes

All files are designed to work together when saved in the same folder. The macros use `ThisWorkbook.Path` to locate sibling files, so no manual path configuration is required as long as all five files stay together.

## Known Limitations

The AP Log's Power Query connections (Inventory Inflow and Inventory_AP_Match) reference a fixed file path to `Inflow_Staging.xlsx`. In production, this pointed to a shared OneDrive folder so the Inventory team and Accounts Payable team could each refresh from the same staging file in real time. To refresh this query against a local copy of `Inflow_Staging.xlsx`, open Power Query Editor (Data → Queries & Connections → right-click the query → Edit) and update the file path in the Source step.

## Skills Demonstrated

Excel VBA, UserForms, FIFO Inventory Costing, Power Query, Cross-Workbook Automation, Finance Approval Workflows, Variance Analysis, AP Reconciliation

---

All data is synthetically generated for portfolio demonstration purposes.

**Analyst:** Shaniah Edwards | MBA, Business Analytics
