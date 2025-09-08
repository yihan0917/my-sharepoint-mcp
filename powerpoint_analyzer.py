#!/usr/bin/env python3
"""
PowerPoint Analyzer for SharePoint Files
Analyzes PowerPoint files on SharePoint by extracting text content
"""

import os
import sys
import requests
from io import BytesIO
from datetime import datetime
import zipfile
import xml.etree.ElementTree as ET
import re
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

async def find_powerpoint_file(filename_pattern, context):
    """Find PowerPoint file on SharePoint by name pattern"""
    print(f"Searching for PowerPoint file matching: '{filename_pattern}'")
    
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
        
        # Search for PowerPoint files in all drives
        pptx_files = []
        for drive in drives:
            drive_id = drive["id"]
            drive_name = drive["name"]
            print(f"Searching in library: {drive_name}")
            
            try:
                # Get root folder contents
                items_response = await graph_client.list_document_contents(site_id, drive_id, "root")
                items = items_response.get("value", [])
                
                # Search through folders and files
                await _search_items_for_powerpoint(graph_client, site_id, drive_id, items, filename_pattern, pptx_files)
                
            except Exception as e:
                print(f"Error searching drive {drive_name}: {str(e)}")
                continue
        
        if not pptx_files:
            print(f"No PowerPoint files found matching '{filename_pattern}'")
            return None
        
        # Find best match
        best_match = _find_best_match(pptx_files, filename_pattern)
        print(f"Best match found: {best_match['filename']}")
        
        return best_match
        
    except Exception as e:
        print(f"Error in PowerPoint file search: {str(e)}")
        return None

async def _search_items_for_powerpoint(graph_client, site_id, drive_id, items, pattern, pptx_files):
    """Recursively search items for PowerPoint files"""
    for item in items:
        if item.get("file") and item["name"].endswith(('.pptx', '.ppt')):
            # Check if file matches pattern
            if _matches_pattern(item["name"], pattern):
                pptx_files.append({
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
                await _search_items_for_powerpoint(graph_client, site_id, drive_id, folder_items, pattern, pptx_files)
            except Exception as e:
                print(f"Error searching folder {item['name']}: {str(e)}")
                continue

def _matches_pattern(filename, pattern):
    """Check if filename matches the search pattern"""
    filename_lower = filename.lower()
    pattern_lower = pattern.lower()
    
    # Remove quotes and extra spaces from pattern
    pattern_clean = pattern_lower.strip().strip("'\"").strip()
    
    # Check for exact substring match first
    if pattern_clean in filename_lower:
        return True
    
    # Extract key terms from pattern
    import re
    pattern_terms = re.findall(r'\b\w{3,}\b', pattern_clean)
    pattern_terms = [term for term in pattern_terms if term not in ['pptx', 'ppt', 'file', 'powerpoint', 'presentation']]
    
    # Check if most key terms are present
    if len(pattern_terms) == 0:
        return False
        
    matches = sum(1 for term in pattern_terms if term in filename_lower)
    match_ratio = matches / len(pattern_terms)
    
    return match_ratio >= 0.4

def _find_best_match(pptx_files, pattern):
    """Find the best matching file from the list"""
    if len(pptx_files) == 1:
        return pptx_files[0]
    
    # Score files based on pattern matching
    scored_files = []
    pattern_lower = pattern.lower()
    
    for file_info in pptx_files:
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
        
        # Prefer more recent files
        if file_info.get("last_modified"):
            score += 1
            
        scored_files.append((score, file_info))
    
    # Return highest scoring file
    scored_files.sort(key=lambda x: x[0], reverse=True)
    return scored_files[0][1]

def extract_text_from_pptx(pptx_data):
    """Extract text content from PowerPoint file"""
    try:
        # PowerPoint files are ZIP archives
        with zipfile.ZipFile(pptx_data, 'r') as zip_file:
            slides_text = []
            slide_files = [f for f in zip_file.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')]
            slide_files.sort()  # Ensure proper order
            
            for slide_file in slide_files:
                try:
                    slide_xml = zip_file.read(slide_file)
                    root = ET.fromstring(slide_xml)
                    
                    # Extract all text elements
                    text_elements = []
                    for elem in root.iter():
                        if elem.tag.endswith('}t'):  # Text elements
                            if elem.text:
                                text_elements.append(elem.text.strip())
                    
                    slide_text = ' '.join(text_elements)
                    if slide_text.strip():
                        slides_text.append(slide_text)
                        
                except Exception as e:
                    print(f"Error processing {slide_file}: {str(e)}")
                    continue
            
            return slides_text
            
    except Exception as e:
        print(f"Error extracting text from PowerPoint: {str(e)}")
        return []

def analyze_hr_metrics(slides_text):
    """Analyze HR-specific metrics from slide text"""
    print("\n" + "="*60)
    print("HR REPORTING ANALYSIS")
    print("="*60)
    
    all_text = ' '.join(slides_text).lower()
    
    # Extract key metrics
    metrics = {}
    
    # Look for hiring numbers
    hire_patterns = [
        r'(\d+)\s*(?:total\s*)?(?:global\s*)?hires?',
        r'hires?[:\s]*(\d+)',
        r'(\d+)\s*hires?\s*(?:globally|total)'
    ]
    
    for pattern in hire_patterns:
        matches = re.findall(pattern, all_text)
        if matches:
            metrics['Total Hires'] = matches[0]
            break
    
    # Look for time to hire
    time_patterns = [
        r'(\d+)\s*days?\s*(?:vs|versus)',
        r'time\s*to\s*hire[:\s]*(\d+)',
        r'(\d+)\s*days?\s*(?:goal|target)'
    ]
    
    for pattern in time_patterns:
        matches = re.findall(pattern, all_text)
        if matches:
            metrics['Time to Hire (days)'] = matches[0]
            break
    
    # Look for acceptance rates
    acceptance_patterns = [
        r'(\d+)%\s*(?:acceptance|offer)',
        r'acceptance[:\s]*(\d+)%'
    ]
    
    for pattern in acceptance_patterns:
        matches = re.findall(pattern, all_text)
        if matches:
            metrics['Offer Acceptance Rate'] = f"{matches[0]}%"
            break
    
    # Look for open positions
    open_patterns = [
        r'(\d+)\s*(?:total\s*)?(?:open|opening)s?',
        r'open\s*positions?[:\s]*(\d+)',
        r'(\d+)\s*positions?\s*open'
    ]
    
    for pattern in open_patterns:
        matches = re.findall(pattern, all_text)
        if matches:
            metrics['Open Positions'] = matches[0]
            break
    
    # Display metrics
    if metrics:
        print("KEY METRICS EXTRACTED:")
        for metric, value in metrics.items():
            print(f"  {metric}: {value}")
    else:
        print("No specific metrics patterns found")
    
    return metrics

async def analyze_powerpoint_file(filename_pattern):
    """Analyze PowerPoint file with detailed function call tracking."""
    
    print("=== POWERPOINT ANALYZER FOR SHAREPOINT ===\n")
    print(f"Target file: {filename_pattern}\n")
    
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
    log_function_call(2, "find_powerpoint_file()", "powerpoint_analyzer.py:31-102", "IN_PROGRESS")
    try:
        file_info = await find_powerpoint_file(filename_pattern, context)
        if not file_info:
            log_function_call(2, "find_powerpoint_file()", "powerpoint_analyzer.py:31-102", "FAILED", "File not found")
            return
        log_function_call(2, "find_powerpoint_file()", "powerpoint_analyzer.py:31-102", "SUCCESS")
        print(f"Found file: {file_info['filename']}")
    except Exception as e:
        log_function_call(2, "find_powerpoint_file()", "powerpoint_analyzer.py:31-102", "FAILED", str(e))
        return
    
    # Step 3: Download PowerPoint File via Graph API
    log_function_call(3, "requests.get() - Graph API download", "Python requests library", "IN_PROGRESS")
    try:
        download_url = f"https://graph.microsoft.com/v1.0/sites/{file_info['site_id']}/drives/{file_info['drive_id']}/items/{file_info['item_id']}/content"
        headers = context.headers.copy()
        headers.pop("Content-Type", None)  # Remove content-type for download
        
        print(f"Download URL: {download_url}")
        response = requests.get(download_url, headers=headers)
        
        if response.status_code == 200:
            log_function_call(3, "requests.get() - Graph API download", "Python requests library", "SUCCESS")
            pptx_data = BytesIO(response.content)
            print(f"Downloaded {len(response.content)} bytes")
        else:
            raise Exception(f"HTTP {response.status_code}: {response.text}")
            
    except Exception as e:
        log_function_call(3, "requests.get() - Graph API download", "Python requests library", "FAILED", str(e))
        return
    
    # Step 4: Extract Text from PowerPoint
    log_function_call(4, "extract_text_from_pptx()", "XML parsing with zipfile", "IN_PROGRESS")
    try:
        slides_text = extract_text_from_pptx(pptx_data)
        log_function_call(4, "extract_text_from_pptx()", "XML parsing with zipfile", "SUCCESS")
        print(f"Extracted text from {len(slides_text)} slides")
    except Exception as e:
        log_function_call(4, "extract_text_from_pptx()", "XML parsing with zipfile", "FAILED", str(e))
        return
    
    # Step 5: Analyze HR Metrics
    log_function_call(5, "analyze_hr_metrics()", "Text analysis and regex", "IN_PROGRESS")
    try:
        metrics = analyze_hr_metrics(slides_text)
        log_function_call(5, "analyze_hr_metrics()", "Text analysis and regex", "SUCCESS")
    except Exception as e:
        log_function_call(5, "analyze_hr_metrics()", "Text analysis and regex", "FAILED", str(e))
        return
    
    # Step 6: Display Slide Content
    log_function_call(6, "Display slide content", "Text processing", "IN_PROGRESS")
    try:
        print("\n" + "="*60)
        print("SLIDE BY SLIDE CONTENT")
        print("="*60)
        
        for i, slide_text in enumerate(slides_text, 1):
            print(f"\n--- SLIDE {i} ---")
            # Clean up text for better readability
            cleaned_text = re.sub(r'\s+', ' ', slide_text).strip()
            if len(cleaned_text) > 500:
                print(cleaned_text[:500] + "...")
            else:
                print(cleaned_text)
        
        log_function_call(6, "Display slide content", "Text processing", "SUCCESS")
        
    except Exception as e:
        log_function_call(6, "Display slide content", "Text processing", "FAILED", str(e))
    
    print("\n" + "="*60)
    print("ANALYSIS COMPLETE")
    print("="*60)
    print("PowerPoint content has been extracted and analyzed.")

def main():
    """Main function with command line argument support"""
    parser = argparse.ArgumentParser(description='Analyze PowerPoint files from SharePoint')
    parser.add_argument('filename', help='Name or pattern of PowerPoint file to analyze')
    
    args = parser.parse_args()
    
    import asyncio
    asyncio.run(analyze_powerpoint_file(args.filename))

if __name__ == "__main__":
    # If no command line args, run interactively
    if len(sys.argv) == 1:
        print("=== Interactive Mode ===")
        filename = input("Enter PowerPoint filename or pattern: ")
        
        import asyncio
        asyncio.run(analyze_powerpoint_file(filename))
    else:
        main()
