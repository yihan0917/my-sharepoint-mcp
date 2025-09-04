#!/usr/bin/env python3
"""
General Excel Analyzer for SharePoint Files
Analyzes any Excel file on SharePoint with function call tracking
"""

import os
import sys
import pandas as pd
import requests
from io import BytesIO
from datetime import datetime
import json
import argparse

# Add the project root to Python path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from auth.sharepoint_auth import get_auth_context

def log_function_call(step, function_name, file_location, status="SUCCESS", error=None):
    """Log each function call with its location and status"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"\n[{timestamp}] Step {step}: {function_name}")
    print(f"File: {file_location}")
    print(f"Status: {status}")
    if error:
        print(f"Error: {error}")
    print("-" * 60)

async def find_excel_file(filename_pattern, context):
    """Find Excel file on SharePoint by name pattern using dynamic discovery"""
    print(f"Searching for Excel file matching: '{filename_pattern}'")
    
    try:
        from utils.graph_client import GraphClient
        from config.settings import SHAREPOINT_CONFIG
        
        # Create Graph client
        graph_client = GraphClient(context)
        
        # Get site info dynamically
        site_parts = SHAREPOINT_CONFIG["site_url"].replace("https://", "").split("/")
        domain = site_parts[0]
        site_name = site_parts[2] if len(site_parts) > 2 else "root"
        site_info = await graph_client.get_site_info(domain, site_name)
        site_id = site_info["id"]
        
        # Get document libraries dynamically
        libraries_response = await graph_client.list_document_libraries(domain, site_name)
        drives = libraries_response.get("value", [])
        
        print(f"Found {len(drives)} document libraries to search")
        
        # Search for Excel files in all drives
        excel_files = []
        for drive in drives:
            drive_id = drive["id"]
            drive_name = drive["name"]
            print(f"Searching in library: {drive_name}")
            
            try:
                # Get root folder contents
                items_response = await graph_client.list_document_contents(site_id, drive_id, "root")
                items = items_response.get("value", [])
                
                # Search through folders and files
                await _search_items_for_excel(graph_client, site_id, drive_id, items, filename_pattern, excel_files)
                
            except Exception as e:
                print(f"Error searching drive {drive_name}: {str(e)}")
                continue
        
        if not excel_files:
            print(f"No Excel files found matching '{filename_pattern}'")
            print("Available Excel files:")
            # List all found Excel files for reference
            all_excel = await _list_all_excel_files(graph_client, site_id, drives)
            for file_info in all_excel[:10]:  # Show first 10
                print(f"  - {file_info['filename']}")
            return None
        
        # Find best match
        best_match = _find_best_match(excel_files, filename_pattern)
        print(f"Best match found: {best_match['filename']}")
        
        return best_match
        
    except Exception as e:
        print(f"Error in dynamic file search: {str(e)}")
        print("Falling back to known file mappings...")
        
        # Fallback to basic site info (remove hardcoded IDs)
        site_id = SHAREPOINT_CONFIG.get("site_id")
        drive_id = SHAREPOINT_CONFIG.get("drive_id")
        
        if not site_id or not drive_id:
            print("No site configuration found. Please check config/settings.py")
            return None
            
        # Basic pattern matching as fallback
        return _fallback_file_search(filename_pattern, site_id, drive_id)

async def _search_items_for_excel(graph_client, site_id, drive_id, items, pattern, excel_files):
    """Recursively search items for Excel files"""
    for item in items:
        if item.get("file") and item["name"].endswith(('.xlsx', '.xls')):
            # Check if file matches pattern
            if _matches_pattern(item["name"], pattern):
                excel_files.append({
                    "site_id": site_id,
                    "drive_id": drive_id,
                    "item_id": item["id"],
                    "filename": item["name"],
                    "size": item.get("size", 0),
                    "last_modified": item.get("lastModifiedDateTime", "")
                })
        
        # Search in folders
        elif item.get("folder"):
            try:
                folder_items_response = await graph_client.list_document_contents(site_id, drive_id, item["id"])
                folder_items = folder_items_response.get("value", [])
                await _search_items_for_excel(graph_client, site_id, drive_id, folder_items, pattern, excel_files)
            except Exception as e:
                print(f"Error searching folder {item['name']}: {str(e)}")
                continue

async def _list_all_excel_files(graph_client, site_id, drives):
    """List all Excel files for reference"""
    all_files = []
    for drive in drives[:3]:  # Limit to first 3 drives
        try:
            items_response = await graph_client.list_document_contents(site_id, drive["id"], "root")
            items = items_response.get("value", [])
            for item in items:
                if item.get("file") and item["name"].endswith(('.xlsx', '.xls')):
                    all_files.append({"filename": item["name"]})
        except:
            continue
    return all_files

def _matches_pattern(filename, pattern):
    """Check if filename matches the search pattern"""
    filename_lower = filename.lower()
    pattern_lower = pattern.lower()
    
    # Remove quotes and extra spaces from pattern
    pattern_clean = pattern_lower.strip().strip("'\"").strip()
    
    # Check for exact substring match first
    if pattern_clean in filename_lower:
        return True
    
    # Extract key terms from pattern (remove common words)
    import re
    # Remove file extensions and common words
    pattern_terms = re.findall(r'\b\w{3,}\b', pattern_clean)
    pattern_terms = [term for term in pattern_terms if term not in ['xlsx', 'file', 'data', 'folder']]
    
    # Check if most key terms are present
    if len(pattern_terms) == 0:
        return False
        
    matches = sum(1 for term in pattern_terms if term in filename_lower)
    match_ratio = matches / len(pattern_terms)
    
    # Lower threshold for better matching
    return match_ratio >= 0.4

def _find_best_match(excel_files, pattern):
    """Find the best matching file from the list"""
    if len(excel_files) == 1:
        return excel_files[0]
    
    # Score files based on pattern matching
    scored_files = []
    pattern_lower = pattern.lower()
    
    for file_info in excel_files:
        filename_lower = file_info["filename"].lower()
        score = 0
        
        # Exact substring match gets highest score
        if pattern_lower in filename_lower:
            score += 100
        
        # Word matches
        pattern_words = pattern_lower.split()
        for word in pattern_words:
            if word in filename_lower:
                score += 10
        
        # Prefer more recent files (if timestamp available)
        if file_info.get("last_modified"):
            score += 1
            
        scored_files.append((score, file_info))
    
    # Return highest scoring file
    scored_files.sort(key=lambda x: x[0], reverse=True)
    return scored_files[0][1]

def _fallback_file_search(filename_pattern, site_id, drive_id):
    """Fallback search using basic pattern matching"""
    print("Using fallback file search...")
    
    # Known file mappings (can be extended)
    known_files = {
        "2023 recruiting dataset": {
            "item_id": "01NL7ATTBNIHMZXNJTZFGJ7A5DL7K5V7AR",
            "filename": "2023 Recruiting Dataset  .xlsx"
        },
        "ngpi metrics": {
            "item_id": "01NL7ATTG5URN4XHBGQFEZCXKYF7ZDW4WM",
            "filename": "NGPI - Metrics Definition & Blockers (in-Progress) - Copy.xlsx"
        }
    }
    
    pattern_lower = filename_pattern.lower()
    for key, file_info in known_files.items():
        if any(word in pattern_lower for word in key.split()):
            return {
                "site_id": site_id,
                "drive_id": drive_id,
                "item_id": file_info["item_id"],
                "filename": file_info["filename"]
            }
    
    print(f"No fallback mapping found for '{filename_pattern}'")
    return None

async def analyze_excel_file(filename_pattern, analysis_type="general"):
    """Analyze any Excel file with detailed function call tracking."""
    
    print("=== GENERAL SHAREPOINT EXCEL ANALYZER ===\n")
    print(f"Target file: {filename_pattern}")
    print(f"Analysis type: {analysis_type}\n")
    
    # Step 1: Get Authentication Context
    log_function_call(1, "get_auth_context()", "/auth/sharepoint_auth.py:214-312", "IN_PROGRESS")
    try:
        context = await get_auth_context()
        log_function_call(1, "get_auth_context()", "/auth/sharepoint_auth.py:214-312", "SUCCESS")
        print(f"Token expires: {context.token_expiry}")
    except Exception as e:
        log_function_call(1, "get_auth_context()", "/auth/sharepoint_auth.py:214-312", "FAILED", str(e))
        return
    
    # Step 2: Find Target File
    log_function_call(2, "find_excel_file()", "general_excel_analyzer.py:31-102", "IN_PROGRESS")
    try:
        file_info = await find_excel_file(filename_pattern, context)
        if not file_info:
            log_function_call(2, "find_excel_file()", "general_excel_analyzer.py:31-102", "FAILED", "File not found")
            return
        log_function_call(2, "find_excel_file()", "general_excel_analyzer.py:31-102", "SUCCESS")
        print(f"Found file: {file_info['filename']}")
    except Exception as e:
        log_function_call(2, "find_excel_file()", "general_excel_analyzer.py:31-102", "FAILED", str(e))
        return
    
    # Step 3: Download Excel File via Graph API
    log_function_call(3, "requests.get() - Graph API download", "Python requests library", "IN_PROGRESS")
    try:
        download_url = f"https://graph.microsoft.com/v1.0/sites/{file_info['site_id']}/drives/{file_info['drive_id']}/items/{file_info['item_id']}/content"
        headers = context.headers.copy()
        headers.pop("Content-Type", None)  # Remove content-type for download
        
        print(f"Download URL: {download_url}")
        response = requests.get(download_url, headers=headers)
        
        if response.status_code == 200:
            log_function_call(3, "requests.get() - Graph API download", "Python requests library", "SUCCESS")
            excel_data = BytesIO(response.content)
            print(f"Downloaded {len(response.content)} bytes")
        else:
            raise Exception(f"HTTP {response.status_code}: {response.text}")
            
    except Exception as e:
        log_function_call(3, "requests.get() - Graph API download", "Python requests library", "FAILED", str(e))
        return
    
    # Step 4: Load Excel with Pandas
    log_function_call(4, "pd.read_excel()", "pandas library", "IN_PROGRESS")
    try:
        # Try to read all sheets first
        df_dict = pd.read_excel(excel_data, sheet_name=None)
        sheet_names = list(df_dict.keys())
        
        # Use first sheet as primary
        df = df_dict[sheet_names[0]]
        
        log_function_call(4, "pd.read_excel()", "pandas library", "SUCCESS")
        print(f"Loaded DataFrame: {df.shape[0]} rows × {df.shape[1]} columns")
        print(f"Available sheets: {sheet_names}")
    except Exception as e:
        log_function_call(4, "pd.read_excel()", "pandas library", "FAILED", str(e))
        return
    
    # Step 5: Basic Data Analysis
    log_function_call(5, "DataFrame analysis methods", "pandas library", "IN_PROGRESS")
    try:
        print("\n" + "="*60)
        print("DATASET OVERVIEW")
        print("="*60)
        print(f"File: {file_info['filename']}")
        print(f"Sheet: {sheet_names[0]}")
        print(f"Shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        
        print(f"\nData Types:")
        for col, dtype in df.dtypes.items():
            print(f"  {col}: {dtype}")
        
        print(f"\nMissing Values:")
        missing = df.isnull().sum()
        for col, count in missing.items():
            if count > 0:
                print(f"  {col}: {count} missing ({count/len(df)*100:.1f}%)")
        
        log_function_call(5, "DataFrame analysis methods", "pandas library", "SUCCESS")
        
    except Exception as e:
        log_function_call(5, "DataFrame analysis methods", "pandas library", "FAILED", str(e))
        return
    
    # Step 6: Analysis Type-Specific Processing
    if analysis_type == "recruiting":
        await analyze_recruiting_metrics(df, 6)
    elif analysis_type == "financial":
        await analyze_financial_metrics(df, 6)
    else:
        await analyze_general_metrics(df, 6)
    
    # Step 7: Sample Data Preview
    log_function_call(7, "df.head() and df.describe()", "pandas display methods", "IN_PROGRESS")
    try:
        print("\n" + "="*60)
        print("SAMPLE DATA (First 5 rows)")
        print("="*60)
        print(df.head())
        
        print("\n" + "="*60)
        print("STATISTICAL SUMMARY")
        print("="*60)
        numeric_cols = df.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0:
            print(df[numeric_cols].describe())
        else:
            print("No numeric columns found for statistical summary")
        
        log_function_call(7, "df.head() and df.describe()", "pandas display methods", "SUCCESS")
        
    except Exception as e:
        log_function_call(7, "df.head() and df.describe()", "pandas display methods", "FAILED", str(e))
    
    print("\n" + "="*60)
    print("ANALYSIS COMPLETE")
    print("="*60)
    print("All function calls and their locations have been documented above.")

async def analyze_recruiting_metrics(df, step_num):
    """Analyze recruiting-specific metrics"""
    log_function_call(step_num, "Recruiting metrics calculation", "pandas aggregation methods", "IN_PROGRESS")
    try:
        print("\n" + "="*60)
        print("RECRUITING METRICS")
        print("="*60)
        
        # Find relevant columns
        metrics = {}
        for col in df.columns:
            col_lower = col.lower()
            if 'application' in col_lower or 'applicant' in col_lower:
                metrics['Total Applications'] = df[col].sum()
            elif 'recruiter' in col_lower and 'screen' in col_lower:
                metrics['Recruiter Screens'] = df[col].sum()
            elif 'hiring' in col_lower and 'screen' in col_lower:
                metrics['Hiring Manager Screens'] = df[col].sum()
            elif 'offer' in col_lower:
                metrics['Total Offers'] = df[col].sum()
            elif 'days' in col_lower and ('open' in col_lower or 'fill' in col_lower):
                metrics['Avg Time to Fill'] = f"{df[col].mean():.1f} days"
                metrics['Time Range'] = f"{df[col].min()}-{df[col].max()} days"
        
        for metric, value in metrics.items():
            print(f"{metric}: {value}")
        
        # Find recruiter performance if applicable
        recruiter_col = None
        for col in df.columns:
            if 'recruiter' in col.lower() and 'screen' not in col.lower():
                recruiter_col = col
                break
        
        if recruiter_col:
            print(f"\nTOP PERFORMERS (by {recruiter_col}):")
            performance = df.groupby(recruiter_col).agg({
                col: 'sum' for col in df.columns 
                if any(word in col.lower() for word in ['application', 'offer', 'screen'])
            })
            print(performance.head(10))
        
        log_function_call(step_num, "Recruiting metrics calculation", "pandas aggregation methods", "SUCCESS")
        
    except Exception as e:
        log_function_call(step_num, "Recruiting metrics calculation", "pandas aggregation methods", "FAILED", str(e))

async def analyze_financial_metrics(df, step_num):
    """Analyze financial-specific metrics"""
    log_function_call(step_num, "Financial metrics calculation", "pandas aggregation methods", "IN_PROGRESS")
    try:
        print("\n" + "="*60)
        print("FINANCIAL METRICS")
        print("="*60)
        
        # Find financial columns
        financial_cols = []
        for col in df.columns:
            col_lower = col.lower()
            if any(word in col_lower for word in ['cost', 'fee', 'budget', 'revenue', 'expense', 'amount']):
                financial_cols.append(col)
        
        if financial_cols:
            for col in financial_cols:
                total = df[col].sum()
                avg = df[col].mean()
                print(f"{col}: Total=${total:,.2f}, Average=${avg:,.2f}")
        else:
            print("No financial columns detected")
        
        log_function_call(step_num, "Financial metrics calculation", "pandas aggregation methods", "SUCCESS")
        
    except Exception as e:
        log_function_call(step_num, "Financial metrics calculation", "pandas aggregation methods", "FAILED", str(e))

async def analyze_general_metrics(df, step_num):
    """Analyze general metrics for any dataset"""
    log_function_call(step_num, "General metrics calculation", "pandas aggregation methods", "IN_PROGRESS")
    try:
        print("\n" + "="*60)
        print("GENERAL METRICS")
        print("="*60)
        
        # Numeric column analysis
        numeric_cols = df.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0:
            print("NUMERIC COLUMNS SUMMARY:")
            for col in numeric_cols:
                print(f"  {col}: Sum={df[col].sum():.2f}, Mean={df[col].mean():.2f}, Max={df[col].max():.2f}")
        
        # Categorical column analysis
        categorical_cols = df.select_dtypes(include=['object']).columns
        if len(categorical_cols) > 0:
            print(f"\nCATEGORICAL COLUMNS:")
            for col in categorical_cols[:5]:  # Limit to first 5
                unique_count = df[col].nunique()
                print(f"  {col}: {unique_count} unique values")
                if unique_count <= 10:
                    print(f"    Values: {list(df[col].unique())}")
        
        log_function_call(step_num, "General metrics calculation", "pandas aggregation methods", "SUCCESS")
        
    except Exception as e:
        log_function_call(step_num, "General metrics calculation", "pandas aggregation methods", "FAILED", str(e))

def main():
    """Main function with command line argument support"""
    parser = argparse.ArgumentParser(description='Analyze Excel files from SharePoint')
    parser.add_argument('filename', help='Name or pattern of Excel file to analyze')
    parser.add_argument('--type', choices=['general', 'recruiting', 'financial'], 
                       default='general', help='Type of analysis to perform')
    
    args = parser.parse_args()
    
    import asyncio
    asyncio.run(analyze_excel_file(args.filename, args.type))

if __name__ == "__main__":
    # If no command line args, run interactively
    if len(sys.argv) == 1:
        print("=== Interactive Mode ===")
        filename = input("Enter Excel filename or pattern: ")
        analysis_type = input("Analysis type (general/recruiting/financial) [general]: ").strip() or "general"
        
        import asyncio
        asyncio.run(analyze_excel_file(filename, analysis_type))
    else:
        main()
