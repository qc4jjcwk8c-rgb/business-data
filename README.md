# Business Data Builder

A web application that helps UK small business owners create structured business planning data through an interactive questionnaire. Generates Excel files with proper formulas, dropdowns, and formatting.

## 🚀 Quick Start

### Local Testing (Easiest)

```bash
# Option 1: Use the start script (Mac/Linux)
./start.sh

# Option 2: Manual start
pip install -r requirements.txt
python3 server.py

# Then open: http://localhost:8000
```

### Deploy to Render.com (Recommended for sharing)

See [GITHUB_DEPLOYMENT.md](GITHUB_DEPLOYMENT.md) for complete deployment instructions.

Quick steps:
1. Push this code to GitHub
2. Connect GitHub to Render.com
3. Deploy as Web Service with `python3 server.py`
4. Get a public URL like `https://your-app.onrender.com`

---

## Features

### Staff Module (Complete)
- Select from 18 predefined team categories + custom teams
- Set employee counts per team with live counter
- Optional additional categorization column
- Generates fully-formatted Excel with:
  - **Dropdowns**: Team selection, PAYE/Contract, OVERHEAD/DIRECT, custom categories
  - **Formulas**: Monthly salary (Annual/12), Total row with SUM formulas
  - **Formatting**: Bold headers, frozen panes (columns A&B, row 1), bold totals
  - **48-month forecast** starting from current month

### Non-Staff Costs Module (Complete)
- Select from 10 cost categories with 70+ predefined items:
  - 🏢 Facilities & Premises
  - 💻 Technology & IT
  - 📦 Equipment & Supplies
  - 🚚 Operations & Delivery Costs
  - 📣 Sales & Marketing
  - 🧾 Professional & External Services
  - 🚗 Travel & Transport
  - 🛡️ Insurance
  - 🏦 Financial & Administrative
  - 🎓 Training & Development
- Adjustable quantities for each item (e.g., 3x "Software subscriptions")
- Live counter showing total line items
- Generates fully-formatted Excel with:
  - **Structure**: Item names with Category and Annual Cost columns
  - **Formulas**: Monthly values (Annual/12), Total row with SUM formulas
  - **Formatting**: Bold headers, frozen panes, bold totals
  - **48-month forecast** starting from current month

### Sales Module (Complete - All 8 Models)
- **8 Business Model Types** with custom questionnaires:
  1. 📦 **Product / Retail / E-commerce**: Units sold, pricing, COGS per unit
  2. 💼 **Service Business**: Hours/projects, rates, delivery costs
  3. 💻 **SaaS / Subscription**: Subscribers, churn rate, growth, per-user costs
  4. 🏪 **Marketplace / Platform**: GMV, commission rates, processing fees
  5. 🏭 **Manufacturing**: Production units, materials, labor, overhead
  6. ⚡ **Usage / Consumption**: Usage metrics, volume, variable costs
  7. 🔀 **Hybrid Models**: Mix multiple revenue streams
  8. ✏️ **Custom Model**: Design your own structure

- **Smart Questionnaires** tailored to each business type
- **Multiple Revenue Streams** - add unlimited items per model
- Generates comprehensive Sales tab with:
  - **REVENUE Section**: All revenue items with growth formulas
  - **COGS Section**: Cost calculations linked to revenue volume
  - **GROSS PROFIT**: Auto-calculated (Revenue - COGS)
  - **Growth Modeling**: Compound growth formulas per item type
  - **48-month projections** with sophisticated formulas

### Output Format
Excel file with 3 tabs (Sales first):
1. **Sales** - Complete revenue & COGS model with gross profit
2. **Staff** - Fully populated with employees and formulas
3. **Non-Staff** - Fully populated with cost items and formulas

## Setup Instructions

### Prerequisites
- Python 3.7+
- Required packages: `openpyxl`

### Installation

1. Install dependencies:
```bash
pip install openpyxl
```

2. Run the server:
```bash
python3 server.py
```

3. Open your browser to:
```
http://localhost:8000
```

## Usage

1. **Start**: Choose "New Model" (upload feature coming soon)
2. **Select Teams**: Pick from predefined categories or add custom ones
3. **Set Employee Counts**: Enter number of employees per team
4. **Add Categories** (Optional): Add an extra categorization column with dropdown options
5. **Select Non-Staff Costs**: Choose items from 10 cost categories, adjust quantities
6. **Choose Sales Model**: Select from 8 business models (Product, Service, SaaS, Marketplace, Manufacturing, Usage, Hybrid, or Custom)
7. **Configure Sales Items**: Answer model-specific questions to build your revenue forecast
8. **Download**: Get your complete business model Excel file with all formatting and formulas

## Excel File Details

### Staff Tab Structure
- **Column A**: Position (e.g., "EM Employee 1")
- **Column B**: Team (dropdown of selected teams)
- **Column C**: Type (dropdown: PAYE or Contract, default PAYE)
- **Column D**: Direct/Overhead (dropdown: OVERHEAD or DIRECT, default OVERHEAD)
- **Column E** (optional): Custom category with user-defined options
- **Column F** (or E if no custom): Annual Salary
- **Columns G+**: Monthly breakdown (48 months from current month)
- **Last Row**: TOTAL with SUM formulas (bold)

### Formula Examples

**Staff Tab:**
- Monthly salary: `=F2/12` (divides annual by 12)
- Total row: `=SUM(G2:G10)` (sums all monthly values)

**Non-Staff Tab:**
- Monthly cost: `=C2/12` (divides annual by 12)
- Total row: `=SUM(D2:D10)` (sums all monthly values)

### Non-Staff Tab Structure
- **Column A**: Item name (e.g., "Rent or lease payments", "Software subscriptions 1", "Software subscriptions 2")
- **Column B**: Category (e.g., "Facilities & Premises", "Technology & IT")
- **Column C**: Annual Cost
- **Columns D+**: Monthly values (48 months from current month) - Formula: `=C2/12`
- **Last Row**: TOTAL with SUM formulas (bold)

### Sales Tab Structure
- **REVENUE Section**:
  - Column A: Item name
  - Column B: Model type (PRODUCT, SAAS, etc.)
  - Column C: Starting value (units, subscribers, GMV, etc.)
  - Column D: Growth rate %
  - Columns E+: Monthly revenue (48 months with compound growth formulas)
  - Total row: SUM of all revenue
- **COGS Section**:
  - Same structure as revenue
  - Costs linked to revenue volume
  - Total row: SUM of all COGS
- **GROSS PROFIT Row**:
  - Formula: Total Revenue - Total COGS
  - Green highlighting
- **Columns D+**: Monthly values (48 months from current month) - Formula: `=C2/12`
- **Last Row**: TOTAL with SUM formulas (bold)

### Formatting
- Headers: Bold, light blue background, centered
- Frozen panes: Columns A & B and Row 1 (Staff), Column A and Row 1 (Non-Staff)
- Total row: Bold text
- All currency: Formatted as `#,##0` (no decimals)
- Column widths optimized for readability

## Deployment to GitHub

1. Create a new repository on GitHub
2. Push these files:
```bash
git init
git add server.py index.html README.md
git commit -m "Initial commit: Business Data Builder"
git remote add origin https://github.com/yourusername/business-data-builder.git
git push -u origin main
```

3. For GitHub Pages (static hosting):
   - Note: The current version requires Python backend for full Excel generation
   - Consider deploying to a service that supports Python (Heroku, PythonAnywhere, etc.)

## Technical Details

### Technologies Used
- **Frontend**: React (via CDN), vanilla JavaScript
- **Backend**: Python HTTP server
- **Excel Generation**: openpyxl library
- **Design**: Custom CSS with brand colors

### Browser Compatibility
- Modern browsers (Chrome, Firefox, Safari, Edge)
- Mobile responsive design

## Future Enhancements

1. **Income Statement Upload**: Parse uploaded P&L to extract line items
2. **Advanced Sales Models**: Industry-specific templates (e.g., franchises, multi-location retail)
3. **Scenario Planning**: Build multiple forecast scenarios (optimistic, realistic, conservative)
4. **Database Storage**: Save and load previous models
5. **Cloud Hosting**: Deploy to cloud platform
6. **Direct Sherloc Integration**: Upload directly to Sherloc forecasting platform

## Support

Created by Jeff Lucas
