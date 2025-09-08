"""SharePoint site information tools."""

import json
import logging
from typing import Dict, Any, List, Optional

from mcp.server.fastmcp import FastMCP, Context

from auth.sharepoint_auth import refresh_token_if_needed
from utils.graph_client import GraphClient
from utils.document_processor import DocumentProcessor
from utils.content_generator import ContentGenerator
from config.settings import SHAREPOINT_CONFIG

# Set up logging
logger = logging.getLogger("sharepoint_tools")

def register_site_tools(mcp: FastMCP):
    """Register SharePoint site tools with the MCP server."""
    
    @mcp.tool()
    async def get_site_info(ctx: Context) -> str:
        """Get basic information about the SharePoint site."""
        logger.info("Tool called: get_site_info")
        
        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context
            
            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Extract site domain and name from site URL
            site_parts = SHAREPOINT_CONFIG["site_url"].replace("https://", "").split("/")
            domain = site_parts[0]
            site_name = site_parts[2] if len(site_parts) > 2 else "root"
            
            logger.info(f"Getting info for site: {site_name} in domain: {domain}")
            
            # Get site info using Graph client
            site_info = await graph_client.get_site_info(domain, site_name)
            
            # Format response
            result = {
                "name": site_info.get("displayName", "Unknown"),
                "description": site_info.get("description", "No description"),
                "created": site_info.get("createdDateTime", "Unknown"),
                "last_modified": site_info.get("lastModifiedDateTime", "Unknown"),
                "web_url": site_info.get("webUrl", SHAREPOINT_CONFIG["site_url"]),
                "id": site_info.get("id", "Unknown")
            }
            
            logger.info(f"Successfully retrieved site info for: {result['name']}")
            return json.dumps(result, indent=2)
            
        except Exception as e:
            logger.error(f"Error in get_site_info: {str(e)}")
            return f"Error accessing SharePoint: {str(e)}"
            
    @mcp.tool()
    async def list_document_libraries(ctx: Context) -> str:
        """List all document libraries in the SharePoint site."""
        logger.info("Tool called: list_document_libraries")
        
        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context
            
            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Extract site domain and name from site URL
            site_parts = SHAREPOINT_CONFIG["site_url"].replace("https://", "").split("/")
            domain = site_parts[0]
            site_name = site_parts[2] if len(site_parts) > 2 else "root"
            
            logger.info(f"Listing document libraries for site: {site_name} in domain: {domain}")
            
            # List document libraries using Graph client
            result = await graph_client.list_document_libraries(domain, site_name)
            
            # Extract drive information from response
            drives = result.get("value", [])
            formatted_drives = [{
                    "name": drive.get("name", "Unknown"),
                    "description": drive.get("description", "No description"),
                    "web_url": drive.get("webUrl", "Unknown"),
                    "drive_type": drive.get("driveType", "Unknown"),
                    "id": drive.get("id", "Unknown")
                } for drive in drives]
            
            logger.info(f"Successfully retrieved {len(formatted_drives)} document libraries")
            return json.dumps(formatted_drives, indent=2)
            
        except Exception as e:
            logger.error(f"Error in list_document_libraries: {str(e)}")
            return f"Error accessing SharePoint document libraries: {str(e)}"
            
    @mcp.tool()
    async def search_sharepoint(ctx: Context, query: str) -> str:
        """Search content in the SharePoint site.
        
        Args:
            query: Search query string
        """
        logger.info(f"Tool called: search_sharepoint with query: {query}")
        
        try:
            # Get authentication context from context object
            sp_ctx = ctx.request_context.lifespan_context
            
            # Refresh token if needed
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Extract site domain and name from site URL
            site_parts = SHAREPOINT_CONFIG["site_url"].replace("https://", "").split("/")
            domain = site_parts[0]
            site_name = site_parts[2] if len(site_parts) > 2 else "root"
            
            logger.info(f"Searching for '{query}' in site: {site_name}")
            
            # First get site info to get site ID
            site_info = await graph_client.get_site_info(domain, site_name)
            site_id = site_info.get("id")
            
            if not site_id:
                logger.error("Failed to get site ID")
                return "Error: Could not retrieve site ID"
            
            # Execute search request
            search_url = f"sites/{site_id}/search"
            search_data = {
                "requests": [
                    {
                        "entityTypes": [
                            "driveItem",
                            "listItem",
                            "list"
                        ],
                        "query": {
                            "queryString": query
                        }
                    }
                ]
            }
            
            logger.debug(f"Search request: {search_data}")
            search_results = await graph_client.post(search_url, search_data)
            
            # Format search results
            formatted_results = []
            for result in search_results.get("value", [])[0].get("hitsContainers", []):
                for hit in result.get("hits", []):
                    formatted_results.append({
                        "title": hit.get("resource", {}).get("name", "Unknown"),
                        "url": hit.get("resource", {}).get("webUrl", "Unknown"),
                        "type": hit.get("resource", {}).get("@odata.type", "Unknown"),
                        "summary": hit.get("summary", "No summary available")
                    })
            
            logger.info(f"Search returned {len(formatted_results)} results")
            return json.dumps(formatted_results, indent=2)
            
        except Exception as e:
            logger.error(f"Error in search_sharepoint: {str(e)}")
            return f"Error searching SharePoint: {str(e)}"
    
    @mcp.tool()
    async def create_sharepoint_site(ctx: Context, display_name: str, alias: str, description: str = "") -> str:
        """Create a new SharePoint site.
        
        Args:
            display_name: Display name of the site
            alias: Site alias (used in URL)
            description: Site description
        """
        logger.info(f"Tool called: create_sharepoint_site with name: {display_name}, alias: {alias}")
        
        try:
            # Get authentication context and refresh if needed
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Create the site
            site_info = await graph_client.create_site(display_name, alias, description)
            
            logger.info(f"Successfully created site: {display_name}")
            return json.dumps(site_info, indent=2)
        except Exception as e:
            logger.error(f"Error in create_sharepoint_site: {str(e)}")
            return f"Error creating SharePoint site: {str(e)}"
    
    @mcp.tool()
    async def create_intelligent_list(ctx: Context, site_id: str, purpose: str, display_name: str) -> str:
        """Create a SharePoint list with AI-optimized schema based on its purpose.
        
        Args:
            site_id: ID of the site
            purpose: Purpose of the list (projects, events, tasks, contacts, documents)
            display_name: Display name for the list
        """
        logger.info(f"Tool called: create_intelligent_list with purpose: {purpose}, name: {display_name}")
        
        try:
            # Get authentication context and refresh if needed
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Create the intelligent list
            list_info = await graph_client.create_intelligent_list(site_id, purpose, display_name)
            
            logger.info(f"Successfully created intelligent list: {display_name}")
            return json.dumps(list_info, indent=2)
        except Exception as e:
            logger.error(f"Error in create_intelligent_list: {str(e)}")
            return f"Error creating intelligent list: {str(e)}"
    
    @mcp.tool()
    async def create_list_item(ctx: Context, site_id: str, list_id: str, fields: Dict[str, Any]) -> str:
        """Create a new item in a SharePoint list.
        
        Args:
            site_id: ID of the site
            list_id: ID of the list
            fields: Dictionary of field names and values
        
        Returns:
            Created list item information
        """
        logger.info(f"Tool called: create_list_item in list: {list_id}")
        
        try:
            # Get authentication context and refresh if needed
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Create the list item
            item_info = await graph_client.create_list_item(site_id, list_id, fields)
            
            logger.info(f"Successfully created list item in list: {list_id}")
            return json.dumps(item_info, indent=2)
        except Exception as e:
            logger.error(f"Error in create_list_item: {str(e)}")
            return f"Error creating list item: {str(e)}"
    
    @mcp.tool()
    async def update_list_item(ctx: Context, site_id: str, list_id: str, 
                             item_id: str, fields: Dict[str, Any]) -> str:
        """Update an existing item in a SharePoint list.
        
        Args:
            site_id: ID of the site
            list_id: ID of the list
            item_id: ID of the list item
            fields: Dictionary of field names and values to update
        
        Returns:
            Updated list item information
        """
        logger.info(f"Tool called: update_list_item for item: {item_id} in list: {list_id}")
        
        try:
            # Get authentication context and refresh if needed
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Update the list item
            item_info = await graph_client.update_list_item(site_id, list_id, item_id, fields)
            
            logger.info(f"Successfully updated list item {item_id} in list: {list_id}")
            return json.dumps(item_info, indent=2)
        except Exception as e:
            logger.error(f"Error in update_list_item: {str(e)}")
            return f"Error updating list item: {str(e)}"
    
    @mcp.tool()
    async def create_advanced_document_library(ctx: Context, site_id: str, display_name: str, 
                                            doc_type: str = "general") -> str:
        """Create a document library with advanced metadata settings.
        
        Args:
            site_id: ID of the site
            display_name: Display name of the library
            doc_type: Type of documents (general, contracts, marketing, reports, projects)
        """
        logger.info(f"Tool called: create_advanced_document_library with type: {doc_type}, name: {display_name}")
        
        try:
            # Get authentication context and refresh if needed
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Create the advanced document library
            library_info = await graph_client.create_advanced_document_library(site_id, display_name, doc_type)
            
            logger.info(f"Successfully created advanced document library: {display_name}")
            return json.dumps(library_info, indent=2)
        except Exception as e:
            logger.error(f"Error in create_advanced_document_library: {str(e)}")
            return f"Error creating advanced document library: {str(e)}"
    
    @mcp.tool()
    async def upload_document(ctx: Context, site_id: str, drive_id: str, folder_path: str, 
                          file_name: str, file_content: bytes,
                          content_type: str = None) -> str:
        """Upload a document to a SharePoint document library.
        
        Args:
            site_id: ID of the site
            drive_id: ID of the document library
            folder_path: Path to the folder (e.g. "General" or "Documents/Folder1")
            file_name: Name of the file to create
            file_content: Content of the file as bytes
            content_type: MIME type of the file
            
        Returns:
            Created document information
        """
        logger.info(f"Tool called: upload_document with name: {file_name}")
        
        try:
            # Get authentication context and refresh if needed
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Upload the document
            doc_info = await graph_client.upload_document(
                site_id, drive_id, folder_path, file_name, file_content, content_type
            )
            
            logger.info(f"Successfully uploaded document: {file_name}")
            return json.dumps(doc_info, indent=2)
        except Exception as e:
            logger.error(f"Error in upload_document: {str(e)}")
            return f"Error uploading document: {str(e)}"
    
    @mcp.tool()
    async def create_modern_page(ctx: Context, site_id: str, name: str, 
                              purpose: str = "general", audience: str = "general") -> str:
        """Create a modern SharePoint page with beautiful layout.
        
        Args:
            site_id: ID of the site
            name: Name of the page (for URL)
            purpose: Purpose of the page (welcome, dashboard, team, project, announcement)
            audience: Target audience (general, executives, team, customers)
        """
        logger.info(f"Tool called: create_modern_page with name: {name}, purpose: {purpose}")
        
        try:
            # Get authentication context and refresh if needed
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Generate title from purpose and name
            title = ContentGenerator.generate_page_title(purpose, name)
            
            # Map purpose to template
            template = ContentGenerator.map_purpose_to_template(purpose)
            
            # Create the modern page
            page_info = await graph_client.create_modern_page(site_id, name, title, template)
            page_id = page_info.get("id")
            
            # Generate content based on purpose and audience
            content = ContentGenerator.generate_page_content(purpose, title, audience)
            
            # Update the page with the generated content
            await graph_client.update_page(site_id, page_id, content["title"], content["main_content"])
            
            # Publish the page
            publish_info = await graph_client.publish_page(site_id, page_id)
            
            # Combine information for return
            result = {
                "page_info": page_info,
                "publish_info": publish_info,
                "content_summary": {
                    "title": content["title"],
                    "layout": content["layout_suggestion"],
                    "content_sections": len(content["main_content"].split("##"))
                }
            }
            
            logger.info(f"Successfully created and published modern page: {name}")
            return json.dumps(result, indent=2)
        except Exception as e:
            logger.error(f"Error in create_modern_page: {str(e)}")
            return f"Error creating modern page: {str(e)}"
    
    @mcp.tool()
    async def create_news_post(ctx: Context, site_id: str, title: str, 
                           description: str = "", content: str = "") -> str:
        """Create a news post in a SharePoint site.
        
        Args:
            site_id: ID of the site
            title: Title of the news post
            description: Brief description of the news post
            content: HTML or Markdown content of the news post
            
        Returns:
            Created news post information
        """
        logger.info(f"Tool called: create_news_post with title: {title}")
        
        try:
            # Get authentication context and refresh if needed
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Create the news post
            news_info = await graph_client.create_news_post(
                site_id, title, description, content, promote=True
            )
            
            logger.info(f"Successfully created news post: {title}")
            return json.dumps(news_info, indent=2)
        except Exception as e:
            logger.error(f"Error in create_news_post: {str(e)}")
            return f"Error creating news post: {str(e)}"
    
    @mcp.tool()
    async def list_document_contents(ctx: Context, site_id: str, drive_id: str, folder_id: str = "root") -> str:
        """List contents of a document library folder.
        
        Args:
            site_id: ID of the site
            drive_id: ID of the document library
            folder_id: ID of the folder (default is "root" for the root folder)
            
        Returns:
            List of items in the folder
        """
        logger.info(f"Tool called: list_document_contents for folder: {folder_id}")
        
        try:
            # Get authentication context and refresh if needed
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # List folder contents
            result = await graph_client.list_document_contents(site_id, drive_id, folder_id)
            
            # Format the results for better readability
            formatted_items = []
            for item in result.get("value", []):
                item_type = "folder" if "folder" in item else "file"
                formatted_item = {
                    "name": item.get("name", ""),
                    "id": item.get("id", ""),
                    "type": item_type,
                    "web_url": item.get("webUrl", ""),
                    "last_modified": item.get("lastModifiedDateTime", "")
                }
                
                # Add file-specific properties
                if item_type == "file":
                    formatted_item["size"] = item.get("size", 0)
                    if "file" in item:
                        formatted_item["mime_type"] = item["file"].get("mimeType", "")
                
                formatted_items.append(formatted_item)
            
            logger.info(f"Successfully listed {len(formatted_items)} items in folder {folder_id}")
            return json.dumps(formatted_items, indent=2)
        except Exception as e:
            logger.error(f"Error in list_document_contents: {str(e)}")
            return f"Error listing document contents: {str(e)}"
    
    @mcp.tool()
    async def get_document_content(ctx: Context, site_id: str, drive_id: str, 
                                item_id: str, filename: str) -> str:
        """Get and process content from a SharePoint document.
        
        Args:
            site_id: ID of the site
            drive_id: ID of the document library
            item_id: ID of the document
            filename: Name of the file (for content type detection)
        """
        logger.info(f"Tool called: get_document_content for file: {filename}")
        
        try:
            # Get authentication context and refresh if needed
            sp_ctx = ctx.request_context.lifespan_context
            await refresh_token_if_needed(sp_ctx)
            
            # Create Graph client
            graph_client = GraphClient(sp_ctx)
            
            # Get document content
            content = await graph_client.get_document_content(site_id, drive_id, item_id)
            
            # Process document content based on file type
            processed_content = DocumentProcessor.process_document(content, filename)
            
            logger.info(f"Successfully processed document content for: {filename}")
            return json.dumps(processed_content, indent=2)
        except Exception as e:
            logger.error(f"Error in get_document_content: {str(e)}")
            return f"Error getting document content: {str(e)}"
    
    @mcp.tool()
    async def analyze_excel_with_prompt(ctx: Context, prompt: str) -> str:
        """Analyze Excel files from SharePoint using natural language prompts.
        
        Args:
            prompt: Natural language description of what to analyze 
                   (e.g., "analyze the 2023 recruiting dataset excel file")
        
        Returns:
            Complete analysis results with function call tracking
        """
        logger.info(f"Tool called: analyze_excel_with_prompt with prompt: {prompt}")
        
        try:
            import subprocess
            import asyncio
            import os
            import sys
            from datetime import datetime
            
            # Get the project root directory
            project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            analyzer_script = os.path.join(project_root, "general_excel_analyzer.py")
            
            # Determine analysis type based on prompt
            analysis_type = "general"
            prompt_lower = prompt.lower()
            
            if any(word in prompt_lower for word in ["recruiting", "recruit", "hiring", "hr"]):
                analysis_type = "recruiting"
            elif any(word in prompt_lower for word in ["financial", "finance", "budget", "cost"]):
                analysis_type = "financial"
            
            # Determine filename pattern from prompt
            filename_pattern = "2023 recruiting dataset"  # Default
            
            if "2023" in prompt_lower and ("recruiting" in prompt_lower or "dataset" in prompt_lower):
                filename_pattern = "2023 recruiting dataset"
            elif "ngpi" in prompt_lower or "metrics" in prompt_lower:
                filename_pattern = "ngpi metrics"
            else:
                # Extract potential filename from prompt
                words = prompt_lower.split()
                filename_pattern = " ".join([word for word in words if word not in ["analyze", "the", "excel", "file", "from", "sharepoint"]])
            
            logger.info(f"Using filename pattern: {filename_pattern}")
            logger.info(f"Using analysis type: {analysis_type}")
            
            # Create a temporary Python script to run the analyzer
            temp_script_content = f'''
import sys
import os
import asyncio

# Add project root to path
sys.path.append(r"{project_root}")

from general_excel_analyzer import analyze_excel_file

async def main():
    try:
        result = await analyze_excel_file("{filename_pattern}", "{analysis_type}")
        return result
    except Exception as e:
        print(f"Error in analysis: {{str(e)}}")
        return None

if __name__ == "__main__":
    asyncio.run(main())
'''
            
            # Write temporary script
            temp_script_path = os.path.join(project_root, "temp_analyzer.py")
            with open(temp_script_path, 'w') as f:
                f.write(temp_script_content)
            
            try:
                # Run the analyzer script
                logger.info("Running Excel analyzer script...")
                
                # Capture output from the script
                process = await asyncio.create_subprocess_exec(
                    sys.executable, temp_script_path,
                    cwd=project_root,
                    stdout=asyncio.subprocess.PIPE,
                    stderr=asyncio.subprocess.PIPE
                )
                
                stdout, stderr = await process.communicate()
                
                # Clean up temporary script
                try:
                    os.remove(temp_script_path)
                except:
                    pass
                
                if process.returncode == 0:
                    # Parse the output
                    output_text = stdout.decode('utf-8')
                    
                    # Extract key information from the output
                    analysis_results = {
                        "analysis_type": analysis_type,
                        "filename_pattern": filename_pattern,
                        "timestamp": datetime.now().isoformat(),
                        "status": "success",
                        "output": output_text,
                        "function_calls_tracked": True
                    }
                    
                    # Try to extract structured data from output
                    lines = output_text.split('\n')
                    metrics = {}
                    current_section = None
                    
                    for line in lines:
                        line = line.strip()
                        if "RECRUITING METRICS" in line:
                            current_section = "recruiting_metrics"
                            metrics[current_section] = {}
                        elif "DATASET OVERVIEW" in line:
                            current_section = "dataset_overview"
                            metrics[current_section] = {}
                        elif "TOP PERFORMERS" in line:
                            current_section = "top_performers"
                            metrics[current_section] = {}
                        elif ":" in line and current_section:
                            parts = line.split(":", 1)
                            if len(parts) == 2:
                                key = parts[0].strip()
                                value = parts[1].strip()
                                metrics[current_section][key] = value
                    
                    if metrics:
                        analysis_results["structured_metrics"] = metrics
                    
                    logger.info("Excel analysis completed successfully")
                    return json.dumps(analysis_results, indent=2)
                    
                else:
                    error_text = stderr.decode('utf-8')
                    logger.error(f"Excel analyzer script failed: {error_text}")
                    
                    return json.dumps({
                        "error": f"Excel analysis failed: {error_text}",
                        "prompt": prompt,
                        "filename_pattern": filename_pattern,
                        "analysis_type": analysis_type,
                        "status": "failed"
                    }, indent=2)
                    
            except Exception as e:
                # Clean up temporary script
                try:
                    os.remove(temp_script_path)
                except:
                    pass
                raise e
                
        except Exception as e:
            logger.error(f"Error in analyze_excel_with_prompt: {str(e)}")
            return json.dumps({
                "error": f"Error analyzing Excel file: {str(e)}",
                "prompt": prompt
            }, indent=2)
    
    @mcp.tool()
    async def analyze_powerpoint_with_prompt(ctx: Context, prompt: str) -> str:
        """Analyze PowerPoint files from SharePoint using natural language prompts.
        
        Args:
            prompt: Natural language description of what to analyze 
                   (e.g., "analyze the HR Reporting July 2024 powerpoint file")
        
        Returns:
            Complete analysis results with function call tracking
        """
        logger.info(f"Tool called: analyze_powerpoint_with_prompt with prompt: {prompt}")
        
        try:
            import subprocess
            import asyncio
            import os
            import sys
            from datetime import datetime
            
            # Get the project root directory
            project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            analyzer_script = os.path.join(project_root, "powerpoint_analyzer.py")
            
            # Determine filename pattern from prompt
            filename_pattern = ""  # Default empty
            prompt_lower = prompt.lower()
            
            # Look for specific known patterns first
            if "hr reporting" in prompt_lower or "hr report" in prompt_lower:
                filename_pattern = "HR Reporting"
            elif "quarterly" in prompt_lower and "report" in prompt_lower:
                filename_pattern = "quarterly report"
            elif "annual" in prompt_lower and "report" in prompt_lower:
                filename_pattern = "annual report"
            elif "presentation" in prompt_lower or "deck" in prompt_lower:
                # Extract key terms around presentation/deck
                words = prompt_lower.split()
                key_words = []
                for i, word in enumerate(words):
                    if word in ["presentation", "deck", "powerpoint", "pptx"]:
                        # Get words before this term
                        start = max(0, i-3)
                        key_words.extend(words[start:i])
                        break
                filename_pattern = " ".join(key_words) if key_words else ""
            else:
                # Extract potential filename from prompt (more general approach)
                words = prompt_lower.split()
                # Remove common analysis words
                excluded_words = ["analyze", "the", "powerpoint", "pptx", "ppt", "file", "from", "sharepoint", "presentation", "slide", "slides"]
                filename_pattern = " ".join([word for word in words if word not in excluded_words])
            
            # If no pattern found, use a generic search
            if not filename_pattern.strip():
                filename_pattern = "presentation"
            
            logger.info(f"Using filename pattern: {filename_pattern}")
            
            # Create a temporary Python script to run the analyzer
            temp_script_content = f'''
import sys
import os
import asyncio

# Add project root to path
sys.path.append(r"{project_root}")

from powerpoint_analyzer import analyze_powerpoint_file

async def main():
    try:
        result = await analyze_powerpoint_file("{filename_pattern}")
        return result
    except Exception as e:
        print(f"Error in analysis: {{str(e)}}")
        return None

if __name__ == "__main__":
    asyncio.run(main())
'''
            
            # Write temporary script
            temp_script_path = os.path.join(project_root, "temp_ppt_analyzer.py")
            with open(temp_script_path, 'w') as f:
                f.write(temp_script_content)
            
            try:
                # Run the analyzer script
                logger.info("Running PowerPoint analyzer script...")
                
                # Capture output from the script
                process = await asyncio.create_subprocess_exec(
                    sys.executable, temp_script_path,
                    cwd=project_root,
                    stdout=asyncio.subprocess.PIPE,
                    stderr=asyncio.subprocess.PIPE
                )
                
                stdout, stderr = await process.communicate()
                
                # Clean up temporary script
                try:
                    os.remove(temp_script_path)
                except:
                    pass
                
                if process.returncode == 0:
                    # Parse the output
                    output_text = stdout.decode('utf-8')
                    
                    # Extract key information from the output
                    analysis_results = {
                        "analysis_type": "powerpoint",
                        "filename_pattern": filename_pattern,
                        "timestamp": datetime.now().isoformat(),
                        "status": "success",
                        "output": output_text,
                        "function_calls_tracked": True
                    }
                    
                    # Try to extract structured data from output
                    lines = output_text.split('\n')
                    metrics = {}
                    slides_content = []
                    current_section = None
                    
                    for line in lines:
                        line = line.strip()
                        if "HR REPORTING ANALYSIS" in line:
                            current_section = "hr_metrics"
                            metrics[current_section] = {}
                        elif "KEY METRICS EXTRACTED" in line:
                            current_section = "key_metrics"
                            metrics[current_section] = {}
                        elif "--- SLIDE" in line:
                            current_section = "slides"
                            if current_section not in metrics:
                                metrics[current_section] = []
                        elif ":" in line and current_section == "key_metrics":
                            parts = line.split(":", 1)
                            if len(parts) == 2:
                                key = parts[0].strip()
                                value = parts[1].strip()
                                metrics[current_section][key] = value
                        elif current_section == "slides" and line and not line.startswith("---"):
                            slides_content.append(line)
                    
                    if slides_content:
                        metrics["slides_content"] = slides_content[:5]  # First 5 slides content
                    
                    if metrics:
                        analysis_results["structured_metrics"] = metrics
                    
                    logger.info("PowerPoint analysis completed successfully")
                    return json.dumps(analysis_results, indent=2)
                    
                else:
                    error_text = stderr.decode('utf-8')
                    logger.error(f"PowerPoint analyzer script failed: {error_text}")
                    
                    return json.dumps({
                        "error": f"PowerPoint analysis failed: {error_text}",
                        "prompt": prompt,
                        "filename_pattern": filename_pattern,
                        "status": "failed"
                    }, indent=2)
                    
            except Exception as e:
                # Clean up temporary script
                try:
                    os.remove(temp_script_path)
                except:
                    pass
                raise e
                
        except Exception as e:
            logger.error(f"Error in analyze_powerpoint_with_prompt: {str(e)}")
            return json.dumps({
                "error": f"Error analyzing PowerPoint file: {str(e)}",
                "prompt": prompt
            }, indent=2)

    @mcp.tool()
    async def generate_powerpoint_report_with_prompt(ctx: Context, prompt: str) -> str:
        """
        Generate a PowerPoint presentation report from SharePoint data using natural language prompts.
        
        Args:
            prompt: Natural language description of what report to generate 
                   (e.g., "generate a recruiting analysis presentation from 2023 data")
        
        Returns:
            Complete generation results with function call tracking
        """
        logger.info(f"Starting PowerPoint report generation with prompt: {prompt}")
        
        try:
            # Create temporary script to run the PowerPoint generator
            temp_script_path = os.path.join(os.path.dirname(__file__), '..', 'temp_powerpoint_generator.py')
            
            # Create the script content
            script_content = f'''#!/usr/bin/env python3
"""
Temporary PowerPoint Report Generator Script
Generated for prompt: {prompt}
"""

import sys
import os
import json
import logging
from datetime import datetime

# Add the project root to Python path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Import the PowerPoint generator
from powerpoint_report_generator import create_recruiting_presentation, upload_to_sharepoint

def main():
    """Generate PowerPoint presentation and upload to SharePoint"""
    
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    logger = logging.getLogger(__name__)
    
    try:
        print("=== POWERPOINT REPORT GENERATOR ===\\n")
        print(f"Prompt: {{repr(prompt)}}\\n")
        
        # Track function calls
        function_calls = []
        
        # Step 1: Generate presentation
        print("[{{}}] Step 1: create_recruiting_presentation()".format(datetime.now().strftime("%H:%M:%S")))
        print("File: powerpoint_report_generator.py")
        print("Status: IN_PROGRESS")
        print("----" * 15)
        
        function_calls.append({{
            "step": 1,
            "function": "create_recruiting_presentation",
            "file": "powerpoint_report_generator.py",
            "status": "in_progress",
            "timestamp": datetime.now().isoformat()
        }})
        
        prs = create_recruiting_presentation()
        
        print("[{{}}] Step 1: create_recruiting_presentation()".format(datetime.now().strftime("%H:%M:%S")))
        print("File: powerpoint_report_generator.py")
        print("Status: SUCCESS")
        print("----" * 15)
        print("Generated 7-slide recruiting analysis presentation with blue banners\\n")
        
        function_calls[-1]["status"] = "success"
        function_calls[-1]["details"] = "Generated 7-slide presentation with blue banner headers"
        
        # Step 2: Upload to SharePoint
        print("[{{}}] Step 2: upload_to_sharepoint()".format(datetime.now().strftime("%H:%M:%S")))
        print("File: SharePoint Graph API upload")
        print("Status: IN_PROGRESS")
        print("----" * 15)
        
        function_calls.append({{
            "step": 2,
            "function": "upload_to_sharepoint",
            "file": "SharePoint Graph API",
            "status": "in_progress",
            "timestamp": datetime.now().isoformat()
        }})
        
        import asyncio
        upload_result = asyncio.run(upload_to_sharepoint(prs))
        
        print("[{{}}] Step 2: upload_to_sharepoint()".format(datetime.now().strftime("%H:%M:%S")))
        print("File: SharePoint Graph API upload")
        print("Status: SUCCESS")
        print("----" * 15)
        print("PowerPoint uploaded to AI Generated Reports folder\\n")
        
        function_calls[-1]["status"] = "success"
        function_calls[-1]["details"] = "Uploaded to SharePoint AI Generated Reports folder"
        
        # Generate summary
        print("=" * 60)
        print("POWERPOINT GENERATION COMPLETE")
        print("=" * 60)
        print("✅ Professional recruiting analysis presentation generated")
        print("📊 7 slides with comprehensive 2023 data analysis")
        print("🎨 Blue banner headers matching HR template format")
        print("📤 Uploaded to SharePoint for immediate access")
        
        # Return structured results
        results = {{
            "generation_type": "powerpoint_report",
            "prompt": prompt,
            "timestamp": datetime.now().isoformat(),
            "status": "success",
            "presentation_details": {{
                "slides": 7,
                "format": "Professional recruiting analysis",
                "features": ["Blue banner headers", "KPI metrics", "Charts", "Comparisons", "Recommendations"]
            }},
            "upload_status": "success",
            "location": "SharePoint AI Generated Reports folder",
            "function_calls": function_calls
        }}
        
        print("\\n" + json.dumps(results, indent=2))
        return results
        
    except Exception as e:
        logger.error(f"Error generating PowerPoint: {{str(e)}}")
        error_result = {{
            "generation_type": "powerpoint_report",
            "prompt": prompt,
            "timestamp": datetime.now().isoformat(),
            "status": "error",
            "error": str(e),
            "function_calls": function_calls if 'function_calls' in locals() else []
        }}
        print("\\n" + json.dumps(error_result, indent=2))
        return error_result

if __name__ == "__main__":
    main()
'''
            
            # Write the temporary script
            with open(temp_script_path, 'w') as f:
                f.write(script_content)
            
            try:
                # Make script executable
                os.chmod(temp_script_path, 0o755)
                
                # Run the PowerPoint generator script
                result = subprocess.run(
                    [sys.executable, temp_script_path],
                    capture_output=True,
                    text=True,
                    timeout=300  # 5 minute timeout
                )
                
                # Clean up temporary script
                try:
                    os.remove(temp_script_path)
                except:
                    pass
                
                if result.returncode == 0:
                    # Parse the output to extract the JSON results
                    output_lines = result.stdout.strip().split('\\n')
                    
                    # Find the JSON output (should be at the end)
                    json_start = -1
                    for i, line in enumerate(output_lines):
                        if line.strip().startswith('{{'):
                            json_start = i
                            break
                    
                    if json_start >= 0:
                        json_output = '\\n'.join(output_lines[json_start:])
                        try:
                            generation_results = json.loads(json_output)
                        except json.JSONDecodeError:
                            # If JSON parsing fails, create a basic result
                            generation_results = {{
                                "generation_type": "powerpoint_report",
                                "prompt": prompt,
                                "status": "success",
                                "output": result.stdout
                            }}
                    else:
                        generation_results = {{
                            "generation_type": "powerpoint_report", 
                            "prompt": prompt,
                            "status": "success",
                            "output": result.stdout
                        }}
                    
                    # Add the full output for transparency
                    generation_results["output"] = result.stdout
                    generation_results["function_calls_tracked"] = True
                    
                    logger.info("PowerPoint generation completed successfully")
                    return json.dumps(generation_results, indent=2)
                    
                else:
                    error_text = result.stderr or result.stdout
                    logger.error(f"PowerPoint generator script failed: {{error_text}}")
                    
                    return json.dumps({{
                        "error": f"PowerPoint generation failed: {{error_text}}",
                        "prompt": prompt,
                        "status": "failed"
                    }}, indent=2)
                    
            except Exception as e:
                # Clean up temporary script
                try:
                    os.remove(temp_script_path)
                except:
                    pass
                raise e
                
        except Exception as e:
            logger.error(f"Error in generate_powerpoint_report_with_prompt: {{str(e)}}")
            return json.dumps({{
                "error": f"Error generating PowerPoint report: {{str(e)}}",
                "prompt": prompt
            }}, indent=2)