# 📊 QBR Dashboard - Interactive Reports on GitHub Pages

A professional, interactive Quarterly Business Review dashboard for SOI tracking and performance metrics. Hosted on GitHub Pages (free) with charts, monthly reports, and real-time data updates.

## ✨ Features

- 📈 **Interactive Charts** - Errors, NAR scores, trends
- 📅 **Monthly Reports** - Detailed breakdown by month
- 👥 **Business Partner Analytics** - Performance by BP
- 📋 **Executive Summary** - Strategic initiatives tracking
- 📊 **Quarterly Views** - Q1-Q4 filtering
- 📱 **Responsive Design** - Works on mobile & desktop
- 🚀 **Hosted on GitHub Pages** - Free, automatic updates

## 🎯 What's Included

```
├── index.html                          # Main dashboard
├── convert-excel.py                   # Excel → JSON converter script
├── GITHUB_SETUP_GUIDE.md             # Detailed setup instructions
├── README.md                          # This file
└── data/                              # JSON data files (generated)
    ├── soi-tracking.json
    ├── soi-completed.json
    └── executive.json
```

## ⚡ Quick Start (5 Minutes)

### 1. Prepare Excel Files
Place your three Excel files in the same folder:
- `SOI_Tracking_Sheet_2026__2_.xlsx`
- `SOI_BPs_Completed_2026.xlsx`
- `Weekly_Executive_Summary.xlsx`

### 2. Convert to JSON
```bash
# Install openpyxl (one-time)
pip install openpyxl

# Run converter
python convert-excel.py
```

This creates a `data/` folder with JSON files.

### 3. Create GitHub Repository
```bash
# Clone the repo you created on GitHub
git clone https://github.com/YOUR_USERNAME/qbr-dashboard.git
cd qbr-dashboard

# Copy all files here (index.html, convert-excel.py, data/ folder, etc.)
```

### 4. Upload to GitHub
```bash
git add .
git commit -m "Initial QBR Dashboard"
git push origin main
```

### 5. Enable GitHub Pages
1. Go to repo **Settings** → **Pages**
2. Select `main` branch as source
3. Save
4. Your dashboard is live! 🎉
   - URL: `https://YOUR_USERNAME.github.io/qbr-dashboard/`

---

## 📊 Dashboard Sections

### Overview Tab
- **KPI Cards**: Total Errors, Errors Worked, Avg NAR Score, Avg Days to Complete
- **Quarterly Charts**: Monthly trends, NAR scores, top BPs by errors
- **Quarter Filter**: Switch between Q1-Q4

### Monthly Reports Tab
- **Monthly Summary**: Total errors, completion rate, avg NAR
- **Detailed Table**: All monthly data points
- **Month Selector**: View any month's data

### Performance Analysis Tab
- **Resolution Rates**: By Business Partner
- **Distribution Charts**: Days to complete analysis
- **Comparative Metrics**: BP performance comparison

### BPs Completed Tab
- **Completion Status**: Tasks completed, scores
- **Status Breakdown**: Done, In Progress, etc.
- **BP Scores**: Individual BP performance

### Executive Summary Tab
- **Strategic Initiatives**: All strategic plans
- **Status Tracking**: On-going, Completed, etc.
- **Timeline View**: Start and end dates

---

## 🔄 Monthly Updates (1 Minute)

To update your dashboard with new data:

```bash
# 1. Update your Excel files
# 2. Run converter
python convert-excel.py

# 3. Commit and push
git add data/
git commit -m "Update monthly data - [Month]"
git push origin main
```

GitHub Pages updates automatically within 30 seconds!

---

## 🎨 Customization

### Change Colors
Edit the CSS gradient in `index.html`:
```css
/* Find this in the <style> section */
background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);

/* Change to your brand colors */
background: linear-gradient(135deg, #YOUR_COLOR_1 0%, #YOUR_COLOR_2 100%);
```

**Color suggestions:**
- Purple theme: `#667eea` → `#764ba2`
- Blue theme: `#0066cc` → `#0088ff`
- Green theme: `#10b981` → `#059669`
- Corporate: `#003366` → `#0066cc`

### Add Custom Tabs
1. Add button in nav:
```html
<button class="nav-btn" data-tab="custom-report">Custom Report</button>
```

2. Add content div:
```html
<div id="custom-report" class="tab-content">
    <!-- Your content here -->
</div>
```

### Change Chart Types
In `index.html`, modify chart creation:
```javascript
updateChart('monthlyErrorsChart', 'line', {  // Change 'line' to:
    // 'bar', 'pie', 'doughnut', 'area', 'radar'
```

---

## 📧 Export Reports

To export dashboard as PDF:
1. Open dashboard in Chrome
2. Press `Ctrl+P` (or `Cmd+P` on Mac)
3. Save as PDF
4. Choose "Save" - creates a snapshot report

---

## 🔧 Troubleshooting

### Dashboard shows "No data"
**Solution**: 
- Verify `data/` folder exists with JSON files
- Check file names match exactly: `soi-tracking.json`, `soi-completed.json`, `executive.json`
- Open browser console (F12) and look for errors

### Charts not displaying
**Solution**:
- Ensure Chart.js CDN is loaded (check Internet connection)
- Verify JSON data format with [jsonlint.com](https://jsonlint.com)
- Check browser console for errors

### GitHub Pages not updating after push
**Solution**:
- Wait 2-3 minutes for GitHub to build
- Hard refresh browser: `Ctrl+Shift+R` (Windows) or `Cmd+Shift+R` (Mac)
- Check repo Settings → Pages to see build status

### Excel conversion fails
**Solution**:
```bash
# Update openpyxl
pip install --upgrade openpyxl

# Check file paths are correct
python convert-excel.py
```

---

## 🚀 Advanced Features

### Automated Monthly Updates
Set up GitHub Actions to automatically update your dashboard:

Create `.github/workflows/auto-update.yml`:
```yaml
name: Auto-Update Data

on:
  schedule:
    - cron: '0 8 1 * *'  # First of every month at 8 AM

jobs:
  update:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v2
      - uses: actions/setup-python@v2
        with:
          python-version: 3.9
      - run: |
          pip install openpyxl
          python convert-excel.py
      - run: |
          git config user.name 'GitHub Actions'
          git config user.email 'actions@github.com'
          git add data/
          git commit -m 'Auto-update: Monthly data'
          git push
```

### Share with Team
Your dashboard is public! Share the URL:
```
https://YOUR_USERNAME.github.io/qbr-dashboard/
```

No passwords needed. Anyone can view real-time data.

### Custom Domain
To use your own domain (e.g., qbr.yourcompany.com):
1. Settings → Pages → Custom domain
2. Add CNAME record to your domain provider
3. GitHub Pages automatically uses your domain

---

## 📋 Checklist

- [ ] Downloaded all three Excel files
- [ ] Created `data/` folder
- [ ] Ran `python convert-excel.py`
- [ ] Created GitHub repository
- [ ] Uploaded files to GitHub
- [ ] Enabled GitHub Pages
- [ ] Verified dashboard loads at HTTPS URL
- [ ] Shared link with team
- [ ] Set up monthly update schedule

---

## 💡 Tips & Tricks

**Tip 1**: Bookmark your dashboard URL for quick access
```
https://YOUR_USERNAME.github.io/qbr-dashboard/
```

**Tip 2**: Use browser's "Add to home screen" on mobile for app-like access

**Tip 3**: Export monthly snapshots to PDF for historical records
- Print → Save as PDF → Organize by month/quarter

**Tip 4**: Share specific quarters with stakeholders
- Add `?quarter=Q3` to URL to deep-link to Q3 data

**Tip 5**: Set calendar reminder for monthly updates
- First day of month: Update Excel files → Run converter → Commit

---

## 📞 Support

If you encounter issues:

1. **Check the detailed guide**: `GITHUB_SETUP_GUIDE.md`
2. **Verify JSON validity**: [jsonlint.com](https://jsonlint.com)
3. **Check browser console**: F12 → Console tab
4. **GitHub Pages status**: Repo → Settings → Pages (shows build errors)

---

## 📈 Next Steps

1. ✅ Set up dashboard
2. ✅ Add to bookmarks
3. ✅ Train team on new tool
4. ✅ Use in QBR meetings
5. ✅ Set up monthly automation
6. ✅ Add to company intranet

---

## 🎉 You're All Set!

Your QBR Dashboard is ready to:
- 📊 Track SOI errors in real-time
- 📈 Monitor BP performance
- 📋 Present executive summaries
- 🎯 Drive data-informed decisions
- 📱 Share across your organization

**Happy tracking! 🚀**

---

*Built with HTML5, Chart.js, and GitHub Pages*
