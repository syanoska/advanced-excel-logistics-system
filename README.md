# Automated Logistics & Inventory Management System (Excel-VBA)

## Overview

This repository features a sophisticated Microsoft Excel-VBA system designed to automate complex logistics, inventory tracking, and dynamic supply chain planning within a demanding operational environment. The solution integrates real-time data entry with advanced Excel formulas and targeted VBA automation to provide comprehensive insights, optimize shipping decisions, and ensure efficient resource allocation.

## Problem Solved

Managing inventory and logistics for a diverse range of items across multiple institutions presents significant challenges, including manual data reconciliation, imprecise forecasting, and time-consuming operational planning. The goal was to develop an intelligent system that transforms fragmented data into a cohesive, automated workflow for accurate inventory assessment, predictive usage analysis, and streamlined delivery scheduling.

## Solution & Key Features

This system is built upon a foundation of interconnected Excel worksheets, enhanced with both powerful formulas and strategic VBA code to deliver a robust and user-friendly experience:

* **'Info Page' - Centralized Configuration Hub:**

    * Serves as the master data source, containing tables for "Sealed Religious Diet Menu" and "Breakfast Menu" items.

    * Includes critical fields like `[Units/mo]` (servings per month per inmate), `[Units/cs]` (servings per case), `[Shipping?]` (dynamic flag 'Y'/'N' to indicate active shipping groups by vendor), and `[Pct]` (for equitable distribution across vendors with varying item counts).

    * This page ensures the system is highly configurable and adaptable to changing operational parameters.

* **'Inv Updates' - Dynamic Data Entry & Visual Aid:**

    * Functions as the primary data entry point for institutional inventories, showing stock levels for every item at every institution.

    * Features an innovative UI/UX enhancement: **dynamic row, column, and cell highlighting** using conditional formatting and a small VBA macro. This visual cue significantly reduces eye strain and improves accuracy when inputting data into large, dense datasets. (This practical solution was inspired by a Excel tips and tricks found on TikTok, showcasing a proactive and resourceful approach to problem-solving).

    * Organizes data using named ranges for structured referencing within formulas.

* **'Usage' - The Core Analytical Engine (Mitochondria of the Workbook):**

    * This sheet is the analytical powerhouse, driven by complex, hidden formulaic calculations, with only two user inputs: `B2` (start date of next cycle) and `C2` (institution selection via data validation).

    * **Automated Delivery Frequency:** Dynamically calculates an institution's delivery frequency (1, 2, or 4 times per 4-week cycle) based on its population, using nested `IF` statements linked to data in 'Inv Updates'.

    * **Detailed Usage Forecasting:** Includes 28 dynamic hidden columns, each representing a day in the cycle. These columns display:

        * Predicted usage between the last inventory update and the next cycle's start date.

        * Real-time daily usage, dynamically categorized by `U` (Usage accounted for/past), `X` (Exact/Current inventory snapshot), and `N` (Needed/Future predicted usage) based on dates and last inventory accurate.

        * This provides a highly granular view of consumption patterns.

    * **Sophisticated Calculations:** Leverages advanced formulas (e.g., `INDEX/MATCH`, `SUM`, `IFERROR`, `IF`, logical functions) to manage complex logic for serving calculations, case conversions, and remaining inventory projections.

    * **Actionable Outputs:** Calculates `Cases Needed` for the next cycle and `Servings in Reserve` at cycle end, providing critical data for procurement and distribution planning.

    * **Interconnectedness:** All three worksheets are seamlessly linked, allowing data to flow from configuration (`Info Page`) and raw inventory (`Inv Updates`) into sophisticated planning and reporting (`Usage`).

* **VBA Automation:** Utilized for enhancing user interaction (the dynamic highlighting in 'Inv Updates').

## Impact & Results

This advanced system significantly enhanced operational efficiency and accuracy:

* **Improved Inventory Accuracy:** Provided a real-time, consolidated view of inventory across all institutions, reducing discrepancies.

* **Optimized Shipping & Distribution:** Enabled precise calculation of cases needed per item per institution, leading to more efficient deliveries and reduced waste.

* **Enhanced Decision-Making:** Offered dynamic forecasting and granular usage insights, allowing management to make proactive, data-driven decisions on stock levels and procurement.

* **Reduced Manual Effort:** Automated complex calculations and data visualization, freeing up significant time previously spent on manual tracking and analysis.

* **Increased User Adoption:** The intuitive UI/UX, including the innovative highlighting feature, made data entry more efficient and less error-prone.

## Technologies Used

* **Microsoft Excel:** Advanced Formulas (`SUMIFS`, `INDEX/MATCH`, `IFERROR`, `IF`, `SUM`, Logical Functions), Excel Tables (ListObjects), Defined Names (Named Ranges), Conditional Formatting, Data Validation, Charts, Pivot Tables.

* **VBA (Visual Basic for Applications):** Event-driven macros for UI enhancement (e.g., dynamic cell highlighting).

## Getting Started / How to Use

1.  **Download & Open:** Clone this repository or download the `excel-vba-logistics-analytics.xlsm` file and open it in Microsoft Excel (ensure macros are enabled).

2.  **Configuration:** Review the **'Info Page'** for menu item definitions, unit conversions, and shipping parameters.

3.  **Update Inventory:** Navigate to the **'Inv Updates'** worksheet and input the latest inventory counts received from institutions. Observe the dynamic highlighting feature as you interact with cells.

4.  **Plan & Analyze:** Go to the **'Usage'** worksheet. Select an institution from the dropdown (`C2`) and specify the beginning date of the next cycle (`B2`) to see real-time forecasts for cases needed and projected reserves.
    *Note: The workbook has been pre-populated with anonymized dummy data to fully convey its functionality.*

## Anonymization Note

Please note that all sensitive and proprietary data from the original project has been replaced with dummy data to protect confidentiality. The structure, formulas, and logic of the system, along with the core VBA functionality, remain fully intact, demonstrating the complete capabilities.

## License

This project is licensed under the MIT License - see the [LICENSE](https://www.google.com/search?q=LICENSE) file for details.
