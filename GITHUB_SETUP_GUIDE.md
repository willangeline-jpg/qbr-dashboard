# QBR Dashboard - GitHub Pages Setup Guide

## Step 1: Create GitHub Repository

1. Go to [github.com/new](https://github.com/new)
2. **Repository name**: `qbr-dashboard`
3. **Description**: "QBR Dashboard - SOI Tracking & Performance Reports"
4. **Public** (to enable GitHub Pages)
5. Click **Create repository**

## Step 2: Clone Repository Locally

```bash
git clone https://github.com/YOUR_USERNAME/qbr-dashboard.git
cd qbr-dashboard
```

## Step 3: Project Structure

Create this folder structure:

```
qbr-dashboard/
├── index.html              # Main dashboard (from the file provided)
├── convert-excel.py        # Data conversion script
├── data/
│   ├── soi-tracking.json   # Generated from SOI Tracking Sheet
│   ├── soi-completed.json  # Generated from BPs Completed
│   └── executive.json      # Generated from Executive Summary
├── assets/
│   ├── style.css          # Optional: extract CSS
│   ├── app.js             # Optional: extract JavaScript
├── README.md              # Documentation
└── .gitignore            # Git ignore file
```

## Step 4: Convert Excel Files to JSON

### Option A: Using Python Script

Create `convert-excel.py`:

```python
import openpyxl
import json
from datetime import datetime

def convert_soi_tracking(filepath):
    """Convert SOI Tracking Sheet to JSON"""
    wb = openpyxl.load_workbook(filepath)
    
    monthly_data = {}
    for month in wb.sheetnames:
        if month.upper() in ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']:
            ws = wb[month]
            rows = list(ws.iter_rows(min_row=3, values_only=True))
            
            month_records = []
            for row in rows[1:]:
                if row[1] and str(row[1]).startswith('AB BP'):
                    try:
                        record = {
                            'bp': str(row[1]),
                            'date': str(row[2]) if row[2] else None,
                            'total_errors': float(row[3]) if row[3] else 0,
                            'errors_worked': float(row[4]) if row[4] else 0,
                            'days_to_complete': float(row[5]) if row[5] else 0,
                            'assigned_to': str(row[6]) if row[6] else None,
                            'nar_score': float(row[14]) if row[14] else 0
                        }
                        month_records.append(record)
                    except Exception as e:
                        pass
            
            if month_records:
                monthly_data[month] = month_records
    
    return monthly_data

def convert_bps_completed(filepath):
    """Convert BPs Completed Sheet to JSON"""
    wb = openpyxl.load_workbook(filepath)
    
    monthly_data = {}
    for month in wb.sheetnames:
        if month.upper() in ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']:
            ws = wb[month]
            rows = list(ws.iter_rows(min_row=4, values_only=True))
            
            month_records = []
            for row in rows:
                if row[0]:  # BP column
                    try:
                        record = {
                            'bp': str(row[0]),
                            'assigned_to': str(row[1]) if row[1] else None,
                            'due_date': str(row[2]) if row[2] else None,
                            'status': str(row[3]) if row[3] else None,
                            'date_completed': str(row[4]) if row[4] else None,
                            'score': float(row[5]) if row[5] else 0
                        }
                        month_records.append(record)
                    except Exception as e:
                        pass
            
            if month_records:
                monthly_data[month] = month_records
    
    return monthly_data

def convert_executive_summary(filepath):
    """Convert Executive Summary to JSON"""
    wb = openpyxl.load_workbook(filepath)
    ws = wb['2026 Summary']
    
    rows = list(ws.iter_rows(min_row=3, values_only=True))
    initiatives = []
    
    for row in rows:
        if row[1]:  # Strategy column
            initiative = {
                'strategy': str(row[1]),
                'goals': str(row[2]) if row[2] else None,
                'actions': str(row[3]) if row[3] else None,
                'owner': str(row[4]) if row[4] else None,
                'status': str(row[5]) if row[5] else None,
                'start_date': str(row[6]) if row[6] else None,
                'end_date': str(row[7]) if row[7] else None,
                'notes': str(row[8]) if row[8] else None
            }
            initiatives.append(initiative)
    
    return initiatives

# Convert files
if __name__ == '__main__':
    print("Converting Excel files...")
    
    soi_tracking = convert_soi_tracking('SOI_Tracking_Sheet_2026__2_.xlsx')
    with open('data/soi-tracking.json', 'w') as f:
        json.dump(soi_tracking, f, indent=2)
    print("✓ SOI Tracking converted")
    
    bps_completed = convert_bps_completed('SOI_BPs_Completed_2026.xlsx')
    with open('data/soi-completed.json', 'w') as f:
        json.dump(bps_completed, f, indent=2)
    print("✓ BPs Completed converted")
    
    executive = convert_executive_summary('Weekly_Executive_Summary.xlsx')
    with open('data/executive.json', 'w') as f:
        json.dump(executive, f, indent=2)
    print("✓ Executive Summary converted")
    
    print("\nAll files converted successfully!")
```

### Option B: Using Online Converter
- Upload Excel to [convertio.co](https://convertio.co/xlsx-json/) or similar service
- Or use Google Sheets → Export as CSV → Convert to JSON

### Run Conversion:
```bash
python convert-excel.py
# Files will be created in data/ folder
```

## Step 5: Update Dashboard to Load Your Data

In `index.html`, replace the mock data section with:

```javascript
// Load actual data from JSON files
let soiTrackingData = {};
let bpsCompletedData = {};
let executiveData = [];

async function loadData() {
    try {
        const tracking = await fetch('data/soi-tracking.json');
        soiTrackingData = await tracking.json();
        
        const completed = await fetch('data/soi-completed.json');
        bpsCompletedData = await completed.json();
        
        const executive = await fetch('data/executive.json');
        executiveData = await executive.json();
        
        console.log('Data loaded successfully');
        updateDashboard();
    } catch (error) {
        console.error('Error loading data:', error);
    }
}

// Call on page load
document.addEventListener('DOMContentLoaded', loadData);
```

## Step 6: Push to GitHub

```bash
git add .
git commit -m "Initial QBR Dashboard setup"
git push origin main
```

## Step 7: Enable GitHub Pages

1. Go to your repo → **Settings** → **Pages**
2. **Source**: Select `main` branch
3. **Folder**: Select `/ (root)`
4. Click **Save**
5. Your site will be live at: `https://YOUR_USERNAME.github.io/qbr-dashboard/`

## Step 8: Update Your Excel Files Monthly

1. Update Excel files locally
2. Run `python convert-excel.py` to regenerate JSON
3. Commit and push:
```bash
git add data/
git commit -m "Update monthly data - September 2026"
git push origin main
```

GitHub Pages will automatically update within seconds!

## Features Included

✅ **Overview Dashboard** - KPI cards and quarter trends
✅ **Monthly Reports** - Detailed data tables by month
✅ **Performance Analysis** - BP-level analytics
✅ **Executive Summary** - Strategic initiatives tracking
✅ **Interactive Charts** - Chart.js integration
✅ **Responsive Design** - Mobile-friendly
✅ **Quarterly Grouping** - Q1-Q4 filtering

## Customization

### Change Colors
In the CSS, modify these variables:
```css
/* Primary gradient */
background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);

/* Change to your brand colors */
background: linear-gradient(135deg, #YOUR_COLOR1 0%, #YOUR_COLOR2 100%);
```

### Add More Tabs
```html
<button class="nav-btn" data-tab="new-tab">New Report</button>
```

```html
<div id="new-tab" class="tab-content">
    <!-- Your content here -->
</div>
```

### Automated Updates
Use GitHub Actions to auto-update data daily:

Create `.github/workflows/update-data.yml`:
```yaml
name: Update Data

on:
  schedule:
    - cron: '0 8 * * MON'  # Every Monday at 8 AM

jobs:
  update:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v2
      - name: Set up Python
        uses: actions/setup-python@v2
        with:
          python-version: '3.x'
      - name: Convert Excel
        run: |
          pip install openpyxl
          python convert-excel.py
      - name: Commit and push
        run: |
          git config user.name "GitHub Actions"
          git config user.email "actions@github.com"
          git add data/
          git commit -m "Auto-update data"
          git push
```

## Troubleshooting

**Dashboard shows "No data"**
- Check that JSON files are in `/data/` folder
- Open browser console (F12) for errors
- Verify JSON format is valid at [jsonlint.com](https://jsonlint.com)

**Charts not displaying**
- Ensure Chart.js CDN link is active
- Check data format matches expected structure

**GitHub Pages not updating**
- Wait 2-3 minutes after push
- Hard refresh browser (Ctrl+Shift+R)
- Check repo Settings → Pages status

## Next Steps

1. ✅ Create repo and push files
2. ✅ Convert Excel to JSON
3. ✅ Enable GitHub Pages
4. ✅ Share your dashboard link: `https://YOUR_USERNAME.github.io/qbr-dashboard/`
5. ✅ Set up monthly update schedule
