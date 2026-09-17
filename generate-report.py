#!/usr/bin/env python3
"""
Automated Report Generator with Graphs
Generates monthly reports with visualizations from Excel data
Runs automatically via GitHub Actions
"""

import openpyxl
import json
import os
from datetime import datetime
from pathlib import Path
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from collections import defaultdict

class ReportGenerator:
    def __init__(self, data_dir='data', reports_dir='reports'):
        self.data_dir = data_dir
        self.reports_dir = reports_dir
        os.makedirs(reports_dir, exist_ok=True)
        os.makedirs(data_dir, exist_ok=True)
        
        # Create graphs subdirectory
        self.graphs_dir = os.path.join(reports_dir, 'graphs')
        os.makedirs(self.graphs_dir, exist_ok=True)
    
    def generate_excel_data(self, soi_tracking_file: str):
        """Convert Excel to JSON and return data"""
        print(f"📊 Reading Excel file: {soi_tracking_file}")
        wb = openpyxl.load_workbook(soi_tracking_file)
        
        monthly_data = {}
        valid_months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
        
        for sheet_name in wb.sheetnames:
            if sheet_name in valid_months:
                ws = wb[sheet_name]
                rows = list(ws.iter_rows(min_row=3, values_only=True))
                month_records = []
                
                for row in rows:
                    if len(row) > 13 and row[1]:
                        try:
                            bp = str(row[1]).strip()
                            if bp.startswith('AB BP'):
                                record = {
                                    'bp': bp,
                                    'date': str(row[2]) if row[2] else None,
                                    'total_errors': float(row[3]) if isinstance(row[3], (int, float)) else 0,
                                    'errors_worked': float(row[4]) if isinstance(row[4], (int, float)) else 0,
                                    'days_to_complete': float(row[5]) if isinstance(row[5], (int, float)) else 0,
                                    'assigned_to': str(row[6]).strip() if row[6] else None,
                                    'nar_score': float(row[14]) if isinstance(row[14], (int, float)) else 0
                                }
                                month_records.append(record)
                        except:
                            pass
                
                if month_records:
                    monthly_data[sheet_name] = month_records
        
        return monthly_data
    
    def create_monthly_report_graphs(self, month: str, data: list):
        """Create graphs for a specific month"""
        print(f"\n📈 Creating graphs for {month}...")
        
        if not data:
            print(f"  No data for {month}")
            return None
        
        # Prepare data for graphs
        total_errors_list = [d['total_errors'] for d in data]
        nar_scores = [d['nar_score'] for d in data]
        days_to_complete = [d['days_to_complete'] for d in data]
        bps = [d['bp'] for d in data]
        
        # Create figure with subplots
        fig, axes = plt.subplots(2, 2, figsize=(14, 10))
        fig.suptitle(f'{month} 2026 - SOI Performance Report', fontsize=16, fontweight='bold')
        
        # Graph 1: Total Errors by Date
        ax1 = axes[0, 0]
        ax1.plot(range(len(total_errors_list)), total_errors_list, marker='o', color='#667eea', linewidth=2)
        ax1.set_title('Total Errors Tracked', fontweight='bold')
        ax1.set_xlabel('Record #')
        ax1.set_ylabel('Number of Errors')
        ax1.grid(True, alpha=0.3)
        ax1.fill_between(range(len(total_errors_list)), total_errors_list, alpha=0.3, color='#667eea')
        
        # Graph 2: NAR Scores
        ax2 = axes[0, 1]
        colors = ['#10b981' if score < 5 else '#f59e0b' if score < 10 else '#ef4444' for score in nar_scores]
        ax2.bar(range(len(nar_scores)), nar_scores, color=colors, alpha=0.7)
        ax2.set_title('NAR Scores (Lower is Better)', fontweight='bold')
        ax2.set_xlabel('Record #')
        ax2.set_ylabel('NAR Score')
        ax2.axhline(y=5, color='orange', linestyle='--', label='Warning', alpha=0.5)
        ax2.legend()
        ax2.grid(True, alpha=0.3, axis='y')
        
        # Graph 3: Days to Complete Distribution
        ax3 = axes[1, 0]
        days_range = [1, 2, 3, 4, 5]
        days_count = {d: days_to_complete.count(d) for d in days_range}
        ax3.bar(days_count.keys(), days_count.values(), color='#764ba2', alpha=0.7)
        ax3.set_title('Resolution Time Distribution', fontweight='bold')
        ax3.set_xlabel('Days to Complete')
        ax3.set_ylabel('Count')
        ax3.grid(True, alpha=0.3, axis='y')
        
        # Graph 4: Top Error BPs
        ax4 = axes[1, 1]
        bp_errors = defaultdict(float)
        for d in data:
            bp_errors[d['bp']] += d['total_errors']
        
        top_bps = sorted(bp_errors.items(), key=lambda x: x[1], reverse=True)[:8]
        bp_names = [bp[0] for bp in top_bps]
        bp_counts = [bp[1] for bp in top_bps]
        
        ax4.barh(bp_names, bp_counts, color='#667eea', alpha=0.7)
        ax4.set_title('Top BPs by Error Count', fontweight='bold')
        ax4.set_xlabel('Total Errors')
        ax4.invert_yaxis()
        ax4.grid(True, alpha=0.3, axis='x')
        
        plt.tight_layout()
        
        # Save the figure
        graph_file = os.path.join(self.graphs_dir, f'{month}_report.png')
        plt.savefig(graph_file, dpi=100, bbox_inches='tight')
        plt.close()
        
        print(f"  ✓ Saved: {graph_file}")
        return graph_file
    
    def create_quarterly_summary_graph(self, monthly_data: dict):
        """Create quarterly summary graphs"""
        print(f"\n📊 Creating quarterly summary graphs...")
        
        quarters = {
            'Q1': ['JAN', 'FEB', 'MAR'],
            'Q2': ['APR', 'MAY', 'JUN'],
            'Q3': ['JUL', 'AUG', 'SEP'],
            'Q4': ['OCT', 'NOV', 'DEC']
        }
        
        fig, axes = plt.subplots(2, 2, figsize=(14, 10))
        fig.suptitle('2026 Quarterly Performance Summary', fontsize=16, fontweight='bold')
        
        # Prepare quarterly data
        quarterly_stats = {}
        for q, months in quarters.items():
            total_errors = 0
            total_worked = 0
            avg_nar = 0
            avg_days = 0
            count = 0
            
            for month in months:
                if month in monthly_data:
                    data = monthly_data[month]
                    total_errors += sum(d['total_errors'] for d in data)
                    total_worked += sum(d['errors_worked'] for d in data)
                    avg_nar += sum(d['nar_score'] for d in data)
                    avg_days += sum(d['days_to_complete'] for d in data)
                    count += len(data)
            
            if count > 0:
                quarterly_stats[q] = {
                    'total_errors': total_errors,
                    'total_worked': total_worked,
                    'avg_nar': avg_nar / count,
                    'avg_days': avg_days / count
                }
        
        # Graph 1: Quarterly Errors
        ax1 = axes[0, 0]
        quarters_list = list(quarterly_stats.keys())
        errors_list = [quarterly_stats[q]['total_errors'] for q in quarters_list]
        ax1.plot(quarters_list, errors_list, marker='o', color='#667eea', linewidth=3, markersize=10)
        ax1.set_title('Total Errors by Quarter', fontweight='bold')
        ax1.set_ylabel('Total Errors')
        ax1.fill_between(range(len(quarters_list)), errors_list, alpha=0.3, color='#667eea')
        ax1.grid(True, alpha=0.3)
        
        # Graph 2: Resolution Rate
        ax2 = axes[0, 1]
        resolution_rates = []
        for q in quarters_list:
            if quarterly_stats[q]['total_errors'] > 0:
                rate = (quarterly_stats[q]['total_worked'] / quarterly_stats[q]['total_errors']) * 100
            else:
                rate = 0
            resolution_rates.append(rate)
        
        bars = ax2.bar(quarters_list, resolution_rates, color=['#10b981', '#3b82f6', '#f59e0b', '#ef4444'])
        ax2.set_title('Error Resolution Rate %', fontweight='bold')
        ax2.set_ylabel('Resolution Rate (%)')
        ax2.set_ylim(0, 100)
        ax2.axhline(y=80, color='green', linestyle='--', label='Target: 80%', alpha=0.5)
        ax2.legend()
        ax2.grid(True, alpha=0.3, axis='y')
        
        # Add value labels on bars
        for bar, rate in zip(bars, resolution_rates):
            height = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width()/2., height,
                    f'{rate:.1f}%', ha='center', va='bottom', fontweight='bold')
        
        # Graph 3: Average NAR Score
        ax3 = axes[1, 0]
        nar_list = [quarterly_stats[q]['avg_nar'] for q in quarters_list]
        ax3.bar(quarters_list, nar_list, color='#764ba2', alpha=0.7)
        ax3.set_title('Average NAR Score by Quarter', fontweight='bold')
        ax3.set_ylabel('Average NAR Score')
        ax3.axhline(y=5, color='orange', linestyle='--', label='Warning Level', alpha=0.5)
        ax3.legend()
        ax3.grid(True, alpha=0.3, axis='y')
        
        # Graph 4: Average Days to Complete
        ax4 = axes[1, 1]
        days_list = [quarterly_stats[q]['avg_days'] for q in quarters_list]
        ax4.plot(quarters_list, days_list, marker='s', color='#f59e0b', linewidth=3, markersize=10)
        ax4.set_title('Average Days to Complete', fontweight='bold')
        ax4.set_ylabel('Days')
        ax4.fill_between(range(len(quarters_list)), days_list, alpha=0.3, color='#f59e0b')
        ax4.grid(True, alpha=0.3)
        
        plt.tight_layout()
        
        # Save the figure
        quarterly_file = os.path.join(self.graphs_dir, 'quarterly_summary.png')
        plt.savefig(quarterly_file, dpi=100, bbox_inches='tight')
        plt.close()
        
        print(f"  ✓ Saved: {quarterly_file}")
        return quarterly_file
    
    def create_html_report(self, monthly_data: dict):
        """Create an HTML report page with all graphs"""
        print(f"\n📄 Creating HTML report page...")
        
        html_content = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>QBR Monthly Report</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
            padding: 40px;
        }
        
        header {
            text-align: center;
            margin-bottom: 40px;
            border-bottom: 3px solid #667eea;
            padding-bottom: 20px;
        }
        
        header h1 {
            font-size: 2.5em;
            color: #333;
            margin-bottom: 10px;
        }
        
        header p {
            font-size: 1.1em;
            color: #666;
        }
        
        .report-section {
            margin-bottom: 40px;
        }
        
        .report-section h2 {
            color: #667eea;
            font-size: 1.8em;
            margin-bottom: 20px;
            border-left: 4px solid #667eea;
            padding-left: 15px;
        }
        
        .report-section img {
            max-width: 100%;
            height: auto;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            margin-bottom: 20px;
        }
        
        .summary-stats {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }
        
        .stat-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
        }
        
        .stat-label {
            font-size: 0.9em;
            opacity: 0.9;
            margin-bottom: 10px;
        }
        
        .stat-value {
            font-size: 2em;
            font-weight: bold;
        }
        
        footer {
            text-align: center;
            margin-top: 40px;
            padding-top: 20px;
            border-top: 1px solid #e0e0e0;
            color: #666;
            font-size: 0.9em;
        }
        
        @media print {
            body {
                background: white;
            }
            .container {
                box-shadow: none;
                padding: 0;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>📊 QBR Performance Report</h1>
            <p>Generated: """ + datetime.now().strftime("%B %d, %Y at %I:%M %p") + """</p>
        </header>
        
        <div class="report-section">
            <h2>📈 Quarterly Summary</h2>
            <img src="graphs/quarterly_summary.png" alt="Quarterly Summary">
        </div>
"""
        
        # Add monthly reports
        months_order = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
        
        for month in months_order:
            if month in monthly_data and os.path.exists(os.path.join(self.graphs_dir, f'{month}_report.png')):
                data = monthly_data[month]
                total_errors = sum(d['total_errors'] for d in data)
                total_worked = sum(d['errors_worked'] for d in data)
                avg_nar = sum(d['nar_score'] for d in data) / len(data) if data else 0
                
                html_content += f"""
        <div class="report-section">
            <h2>{month} 2026 Performance</h2>
            
            <div class="summary-stats">
                <div class="stat-card">
                    <div class="stat-label">Total Errors</div>
                    <div class="stat-value">{int(total_errors)}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Errors Worked</div>
                    <div class="stat-value">{int(total_worked)}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Avg NAR Score</div>
                    <div class="stat-value">{avg_nar:.1f}</div>
                </div>
                <div class="stat-card">
                    <div class="stat-label">Records</div>
                    <div class="stat-value">{len(data)}</div>
                </div>
            </div>
            
            <img src="graphs/{month}_report.png" alt="{month} Report">
        </div>
"""
        
        html_content += """
        <footer>
            <p>This report is auto-generated from SOI tracking data.</p>
            <p>Updated: """ + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + """</p>
        </footer>
    </div>
</body>
</html>
"""
        
        # Save HTML report
        report_file = os.path.join(self.reports_dir, 'index.html')
        with open(report_file, 'w') as f:
            f.write(html_content)
        
        print(f"  ✓ Saved: {report_file}")
        return report_file
    
    def run(self, soi_tracking_file: str):
        """Run the complete report generation"""
        print("\n" + "="*60)
        print("QBR Report Generator - Automated Report Creation")
        print("="*60)
        
        # Generate JSON data from Excel
        monthly_data = self.generate_excel_data(soi_tracking_file)
        
        if not monthly_data:
            print("\n❌ No data found in Excel file!")
            return False
        
        print(f"\n✓ Extracted data for {len(monthly_data)} months")
        
        # Save JSON data
        json_file = os.path.join(self.data_dir, 'soi-tracking.json')
        with open(json_file, 'w') as f:
            json.dump(monthly_data, f, indent=2)
        print(f"✓ Saved JSON: {json_file}")
        
        # Generate monthly graphs
        for month in monthly_data.keys():
            self.create_monthly_report_graphs(month, monthly_data[month])
        
        # Generate quarterly summary
        self.create_quarterly_summary_graph(monthly_data)
        
        # Generate HTML report
        self.create_html_report(monthly_data)
        
        print("\n" + "="*60)
        print("✅ Report generation complete!")
        print("="*60)
        print(f"\nGenerated files:")
        print(f"  📊 HTML Report: reports/index.html")
        print(f"  📈 Quarterly Summary: reports/graphs/quarterly_summary.png")
        print(f"  📋 Monthly Reports: reports/graphs/[MONTH]_report.png")
        print(f"\nAll files ready to commit to GitHub!")
        return True

if __name__ == '__main__':
    # File paths
    SOI_TRACKING_FILE = 'SOI_Tracking_Sheet_2026__2_.xlsx'
    
    generator = ReportGenerator()
    generator.run(SOI_TRACKING_FILE)
