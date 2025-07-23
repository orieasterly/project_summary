# Project Summary

---

## ERP Development & Business Systems (Global Beauty Care)

### ERP Report Bug Fixes  
**Objective:** Resolve critical defects in Acumatica report logic and layout 
**Technologies:** Acumatica Report Designer, Generic Inquiries  
**Contributions:**  
- Identified and resolved issues in logic, formatting, and data binding within report definitions  
- Verified report outputs against source data and business rules to ensure accuracy and compliance  
**Outcome:** Fixed legacies bugs in the report, improved data integrity and accuracy

### Generic Inquiry (GI) Optimization
**Objective:** Design and manage Generic Inquiries in Acumatic to export data and integrate with external reporting and analytics tools  
**Technologies:** Acumatica Generic Inquiries, OData, Power BI  
**Contributions:**
- Developed and refined Generic Inquiries to extract operational data across Sales, Purchasing, Inventory, and Finance  
- Integrated OData endpoints with Excel and Power BI to support live data links, exports, and interactive visualizations  
**Outcome:** Enabled seamless data flow from Acumatica to external applications, improving reporting accuracy and operational visibility

### Import Scenario Automation
**Objective:** Automate and standardize data entry workflows in Acumatica using Import Scenarios  
**Technologies:** Acumatica Import Scenarios, Excel templates  
**Contributions:**
- Configured automation to import SKUs, sales orders, customer profiles, and vendor records 
- Implemented field mapping and validation rules to ensure data consistency and accuracy  
**Outcome:** Minimized manual entry errors and increased efficiency across core data processes

---

## Retail Analytics & Labeling Automation (Global Beauty Care)

### Retail Purchase Order Integration
**Objective:** Extract and structure retailer purchase order data to streamline sales order entry  
**Technologies:** Power BI, Power Query, Excel VBA, Acumatica ERP  
**Contributions:**
- Parsed and normalized multi-store purchase orders from PDFs using Power BI and Power Query  
- Automated store-level file generation and folder structuring via Excel VBA for seamless import into Acumatica  
**Outcome:** Reduced manual PO handling, standardized data preparation, and improved efficiency of sales order processing

### Carton Marking Automation
**Objective:** Auto-generate compliant shipping carton labels for factory use  
**Technologies:** Acumatica OData, Excel, VBA  
**Contributions:**
- Retrieved item data from Acumatica via OData and cross-referenced SKU lists in Excel  
- Built a VBA-driven system to generate formatted carton labels with master and inner case details  
**Outcome:** Increased labeling accuracy and significantly reduced manual preparation effort

### Sales Presentation Automation (Global Beauty Care)
**Objective:** Develop a process for creating branded product slides
**Technologies:** Excel, PowerPoint, VBA  
**Contributions:**
- Mapped product details (images, SKUs, descriptions, pricing) from Excel into PowerPoint templates  
- Developed a VBA tool to auto-generate and format slides consistently  
**Outcome:** Streamlined sales deck production and improved branding consistency

### Product Data & Customer-Specific Catalog Review
**Objective:** Ensure pricing consistency across customer-specific product catalogs  
**Technologies:** Excel OData (Acumatica), Excel VBA  
**Contributions:**
- Pulled live product and pricing data from Acumatica for validation  
- Built a VBA tool to compare pricing across product categories  
**Outcome:** Identified pricing misalignments and delivered actionable insights to sales leadership

### Refactor – Customer-Specific Product Catalogs Macro (Current Project)
**Objective:** Redesign the Excel-based product catalog generator to improve modularity, performance, and scalability  
**Technologies:** Excel VBA (class modules), named ranges, dictionary objects  
**Contributions:**
- Refactored procedural logic into reusable class-based architecture  
- Automated workbook generation by customer and category  
**Outcome:** Improved processing speed and maintainability; supported complex business rules like cross-category handling and customer-specific pricing

### Email Automation with Power Automate
**Objective:** Automate routine email distribution with dynamic file attachments  
**Technologies:** Power Automate, Outlook, OneDrive  
**Contributions:**
- Designed and deployed a Power Automate flow to send emails with context-specific attachments  
**Outcome:** Streamlined recurring communication workflows and reduced manual effort

---

## Financial & Operational Analysis (Premier Beehive, NZ)

### COGS Variance Analysis
**Objective:** Standardize cost of goods sold (COGS) variance methodology  
**Technologies:** Excel, BI tools  
**Contributions:**
- Analyzed cost variances by volume, mix, and rate effects 
- Leveraged production and cost accounting data  
**Outcome:** Improved cost visibility and enhanced budget variance analysis

### SKU Costing Model Automation
**Objective:** Automate SKU-level costing to support pricing decisions  
**Technologies:** Excel, VBA  
**Contributions:**
- Linked external data sources for real-time updates  
- Automated SKU cost calculations including material, labor, and overhead  
**Outcome:** Reduced costing time from days to hours

### Quadrant Analysis for SKU Performance
**Objective:** Categorize product performance using a quadrant-based model  
**Technologies:** Excel, VBA
**Contributions:**
- Defined and visualized performance metrics (volume, margin, growth)  
- Identified top and bottom performers  
**Outcome:** Supported strategic decisions in product optimization and resource planning

---

## Dashboards & Visualization (Premier Beehive, NZ)

### Labor Dashboard
**Objective:** Track labor cost and efficiency in real time  
**Technologies:** Excel, Power Query, Power Pivot, DAX  
**Contributions:**
- Integrated employee, production, and financial data  
- Created dynamic dashboards for operational insights  
**Outcome:** Reduced reporting time and enabled timely decisions

### SKU Listings Dashboard
**Objective:** Analyze SKU profitability across cost components  
**Technologies:** Power BI, Power Query, DAX  
**Contributions:**
- Merged data from sales, labor, production, and COGS  
- Created SKU- and category-level profitability breakdowns  
**Outcome:** Improved financial planning and SKU-level decision-making

### Manufacturing Yield KPI Dashboard
**Objective:** Monitor yield metrics across production stages  
**Technologies:** Power Query, Excel  
**Contributions:**
- Compiled production yield data from input to final output  
- Designed visuals to benchmark performance  
**Outcome:** Identified inefficiencies and improved production quality

### Capital Expenditure Tracking System
**Objective:** Implement real-time CAPEX monitoring tools  
**Technologies:** Power Query, Microsoft BI tools  
**Contributions:**
- Centralized budget, project status, and ROI data  
- Built milestone tracking and automated reporting  
**Outcome:** Enhanced control over capital spending, improved investment tracking, and enabled early detection of project variances to support corrective action
