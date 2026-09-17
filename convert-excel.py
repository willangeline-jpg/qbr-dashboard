#!/usr/bin/env python3
"""
Excel to JSON Converter for QBR Dashboard
Converts XLSX files to JSON format for the GitHub Pages dashboard
"""

import openpyxl
import json
import os
from datetime import datetime
from typing import Dict, List, Any

class ExcelConverter:
    def __init__(self, data_dir='data'):
        self.data_dir = data_dir
        os.makedirs(data_dir, exist_ok=True)
    
    def convert_soi_tracking(self, filepath: str) -> Dict[str, List[Dict]]:
        """
        Convert SOI Tracking Sheet 2026 to JSON
        Extracts monthly sheets and creates monthly data files
        """
        print(f"📊 Processing SOI Tracking Sheet: {filepath}")
        wb = openpyxl.load_workbook(filepath)
        
        monthly_data = {}
        valid_months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
        
        for sheet_name in wb.sheetnames:
            if sheet_name in valid_months:
                print(f"  Processing {sheet_name}...")
                ws = wb[sheet_name]
                
                # Skip header rows and process data
                rows = list(ws.iter_rows(min_row=3, values_only=True))
                month_records = []
                
                for row in rows:
                    # Column layout: B=BP, C=DATE, D=TOTAL_ERRORS, E=ERRORS_WORKED, 
                    #                F=DAYS_TO_COMPLETE, G=ASSIGNED_TO, etc.
                    if len(row) > 13 and row[1]:  # Check if BP column (index 1) has data
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
                        except (ValueError, TypeError, IndexError) as e:
                            continue
                
                if month_records:
                    monthly_data[sheet_name] = month_records
                    print(f"    ✓ {sheet_name}: {len(month_records)} records")
        
        return monthly_data
    
    def convert_bps_completed(self, filepath: str) -> Dict[str, List[Dict]]:
        """
        Convert SOI BPs Completed 2026 to JSON
        Extracts BP completion status and scores
        """
        print(f"✅ Processing BPs Completed: {filepath}")
        wb = openpyxl.load_workbook(filepath)
        
        monthly_data = {}
        valid_months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
        
        for sheet_name in wb.sheetnames:
            if sheet_name in valid_months:
                print(f"  Processing {sheet_name}...")
                ws = wb[sheet_name]
                
                # Start from row 4 (after title rows)
                rows = list(ws.iter_rows(min_row=4, values_only=True))
                month_records = []
                
                for row in rows:
                    # Column layout: A=BP, B=ASSIGNED_TO, C=DUE, D=STATUS, E=DATE_COMPLETED, F=SCORE
                    if len(row) > 5 and row[0]:
                        try:
                            bp = str(row[0]).strip()
                            if bp and not bp.startswith('='):  # Skip formula rows
                                record = {
                                    'bp': bp,
                                    'assigned_to': str(row[1]).strip() if row[1] else None,
                                    'due_date': str(row[2]) if row[2] else None,
                                    'status': str(row[3]).strip() if row[3] else None,
                                    'date_completed': str(row[4]) if row[4] else None,
                                    'score': float(row[5]) if isinstance(row[5], (int, float)) else 0
                                }
                                month_records.append(record)
                        except (ValueError, TypeError, IndexError) as e:
                            continue
                
                if month_records:
                    monthly_data[sheet_name] = month_records
                    print(f"    ✓ {sheet_name}: {len(month_records)} records")
        
        return monthly_data
    
    def convert_executive_summary(self, filepath: str) -> List[Dict]:
        """
        Convert Weekly Executive Summary to JSON
        Extracts strategic initiatives and status
        """
        print(f"📋 Processing Executive Summary: {filepath}")
        wb = openpyxl.load_workbook(filepath)
        
        # Try 2026 Summary first, fall back to others
        sheet_name = '2026 Summary' if '2026 Summary' in wb.sheetnames else wb.sheetnames[1]
        print(f"  Using sheet: {sheet_name}")
        
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(min_row=3, values_only=True))
        initiatives = []
        
        for row in rows:
            # Column layout: B=STRATEGY, C=GOALS, D=ACTIONS, E=OWNER, F=STATUS, 
            #                G=START_DATE, H=END_DATE, I=NOTES
            if len(row) > 8 and row[1]:
                try:
                    initiative = {
                        'strategy': str(row[1]).strip() if row[1] else None,
                        'goals': str(row[2]).strip() if row[2] else None,
                        'actions': str(row[3]).strip() if row[3] else None,
                        'owner': str(row[4]).strip() if row[4] else None,
                        'status': str(row[5]).strip() if row[5] else None,
                        'start_date': str(row[6]) if row[6] else None,
                        'end_date': str(row[7]) if row[7] else None,
                        'notes': str(row[8]).strip() if row[8] else None
                    }
                    if initiative['strategy']:  # Only add if strategy is not empty
                        initiatives.append(initiative)
                except (ValueError, TypeError, IndexError) as e:
                    continue
        
        print(f"  ✓ {len(initiatives)} initiatives")
        return initiatives
    
    def save_json(self, data: Any, filename: str):
        """Save data to JSON file"""
        filepath = os.path.join(self.data_dir, filename)
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        print(f"  💾 Saved to {filepath}")
    
    def run(self, soi_tracking_file: str, bps_completed_file: str, executive_file: str):
        """Run all conversions"""
        print("\n" + "="*60)
        print("QBR Dashboard - Excel to JSON Converter")
        print("="*60 + "\n")
        
        # Convert SOI Tracking
        if os.path.exists(soi_tracking_file):
            soi_data = self.convert_soi_tracking(soi_tracking_file)
            self.save_json(soi_data, 'soi-tracking.json')
        else:
            print(f"❌ File not found: {soi_tracking_file}")
        
        # Convert BPs Completed
        if os.path.exists(bps_completed_file):
            bps_data = self.convert_bps_completed(bps_completed_file)
            self.save_json(bps_data, 'soi-completed.json')
        else:
            print(f"❌ File not found: {bps_completed_file}")
        
        # Convert Executive Summary
        if os.path.exists(executive_file):
            exec_data = self.convert_executive_summary(executive_file)
            self.save_json(exec_data, 'executive.json')
        else:
            print(f"❌ File not found: {executive_file}")
        
        print("\n" + "="*60)
        print("✅ Conversion complete!")
        print("="*60)
        print(f"\nNext steps:")
        print(f"1. Verify JSON files in 'data/' folder")
        print(f"2. Upload to GitHub repository")
        print(f"3. Enable GitHub Pages in repository settings")
        print(f"4. Access your dashboard at: https://YOUR_USERNAME.github.io/qbr-dashboard/")

if __name__ == '__main__':
    # File paths - Update these to match your files
    SOI_TRACKING_FILE = 'SOI_Tracking_Sheet_2026__2_.xlsx'
    BPS_COMPLETED_FILE = 'SOI_BPs_Completed_2026.xlsx'
    EXECUTIVE_FILE = 'Weekly_Executive_Summary.xlsx'
    
    converter = ExcelConverter()
    converter.run(SOI_TRACKING_FILE, BPS_COMPLETED_FILE, EXECUTIVE_FILE)
