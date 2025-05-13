
import os
import re
import sys
import docx
import pandas as pd
from docx.document import Document as DocxDocument
from docx.oxml.shape import CT_Picture
from docx.oxml.ns import qn
from docx.shared import Inches
from PIL import Image
from io import BytesIO
import base64
import zipfile
import json
import logging
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Color

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def extract_tasks_from_word(docx_path, assembly_id, assembly_name, figure_start_range, figure_end_range):
    """
    Extract tasks from a Word document that contains a table.
    
    Args:
        docx_path: Path to the Word document
        assembly_id: Assembly sequence ID to prefix task numbers
        assembly_name: Name of the assembly for task descriptions
        figure_start_range: Starting number for figure references
        figure_end_range: Ending number for figure references
    
    Returns:
        Tuple of (tasks_df, images_dict)
    """
    try:
        # Load the document
        doc = docx.Document(docx_path)
        
        # Initialize variables
        tasks = []
        images = {}
        image_map = {}  # Map to track figure references to image_ids
        image_task_mapping = {}  # Map to track which tasks reference which images
        
        # Find tables in the document
        if len(doc.tables) == 0:
            logger.warning("No tables found in the document")
            return None, {}
        
        # Process each table in the document (assume first table contains tasks)
        table = doc.tables[0]
        
        # Check if first row is likely a header
        header_row = table.rows[0]
        is_header = False
        
        # Check if first row contains header-like text
        header_terms = ['sl no', 'sl.no', 'sl. no', 'serial no', 'task no', 'job details']
        header_cell_texts = [cell.text.lower().strip() for cell in header_row.cells]
        
        for term in header_terms:
            if any(term in text for text in header_cell_texts):
                is_header = True
                logger.info(f"Detected header row: {' | '.join([cell.text for cell in header_row.cells])}")
                break
        
        # Process each row of the table (skipping header if exists)
        for i, row in enumerate(table.rows):
            # Skip if it's the header row we identified
            if i == 0 and is_header:
                continue
                
            # Skip if it's likely another header row
            if any(cell.text.lower().strip() in header_terms for cell in row.cells):
                continue
            
            # Extract data from cells
            try:
                if len(row.cells) < 2:
                    continue  # Skip rows with insufficient cells
                    
                sl_no_cell = row.cells[0].text.strip()
                job_details_cell = row.cells[1]
                
                # Extract the task number
                try:
                    sl_no = int(re.search(r'\d+', sl_no_cell).group())
                except (AttributeError, ValueError):
                    logger.warning(f"Could not parse sl_no from '{sl_no_cell}', skipping row")
                    continue
                
                # Generate task number in the format assembly_id.0.XXX
                task_number = f"{assembly_id}.0.{str(sl_no).zfill(3)}"
                
                # Extract job details text and process paragraphs
                job_details = ""
                image_references = []
                
                # Extract all figure references from the job details text
                figure_pattern = re.compile(r"Figure\s+(\d+)", re.IGNORECASE)
                
                # Process all paragraphs in the job details cell
                for paragraph in job_details_cell.paragraphs:
                    job_details += paragraph.text + "\n"
                    # Look for figure references in this paragraph
                    matches = figure_pattern.findall(paragraph.text)
                    for match in matches:
                        fig_num = int(match)
                        if figure_start_range <= fig_num <= figure_end_range:
                            # Generate image ID for this figure (format: assembly_id-0-XXX)
                            image_id = f"{assembly_id}-0-{str(fig_num).zfill(3)}"
                            
                            # Track which figures are referenced by this task
                            if task_number not in image_task_mapping:
                                image_task_mapping[task_number] = set()
                            
                            image_task_mapping[task_number].add(image_id)
                            
                            # Track figure number to image ID mapping
                            image_map[fig_num] = image_id
                
                # Create task entry with attachment information
                task = {
                    'task_no': task_number,
                    'type': 'Operation',
                    'eta_sec': '',
                    'description': assembly_name,  # Use the provided assembly name as description
                    'activity': job_details.strip(),
                    'specification': '',
                    'attachment': ''  # Will be populated after processing all tasks and images
                }
                
                tasks.append(task)
                
            except Exception as e:
                logger.error(f"Error processing row {i}: {e}")
                continue
        
        # Extract all images from document
        all_rels = []
        
        # Get all the relationships in the document
        for rel in doc.part.rels.values():
            if "image" in rel.target_ref:
                try:
                    image_data = rel.target_part.blob
                    all_rels.append({
                        'id': rel.rId,
                        'data': image_data,
                        'extension': guess_image_extension(image_data)
                    })
                except Exception as e:
                    logger.error(f"Error extracting image: {e}")
        
        logger.info(f"Found {len(all_rels)} images in document relationships")
        logger.info(f"Found {len(image_map)} figure references in text")
        
        # Create final image dictionary with proper IDs
        for fig_num, image_id in image_map.items():
            # Find the corresponding image data
            image_idx = fig_num - figure_start_range
            if 0 <= image_idx < len(all_rels):
                images[image_id] = {
                    'data': all_rels[image_idx]['data'],
                    'extension': all_rels[image_idx]['extension'],
                    'figure_number': fig_num
                }
        
        # If we have more images than figure references, add them with sequential IDs
        for i in range(len(image_map), len(all_rels)):
            next_num = len(image_map) + i + 1
            image_id = f"{assembly_id}-0-{str(next_num).zfill(3)}"
            images[image_id] = {
                'data': all_rels[i]['data'],
                'extension': all_rels[i]['extension'],
                'figure_number': None  # No specific figure number
            }
        
        # Now update each task with its attachment information (comma-separated image IDs)
        for task in tasks:
            task_num = task['task_no']
            if task_num in image_task_mapping:
                task['attachment'] = ', '.join(sorted(list(image_task_mapping[task_num])))
        
        # Create DataFrame
        if tasks:
            df = pd.DataFrame(tasks)
            return df, images
        else:
            logger.warning("No tasks extracted from the document")
            return None, {}
            
    except Exception as e:
        logger.error(f"Error processing document: {e}")
        return None, {}

def guess_image_extension(image_data):
    """
    Determine the image file extension based on its data.
    """
    import imghdr
    img_type = imghdr.what(None, h=image_data)
    if img_type == 'jpeg':
        return 'jpg'
    return img_type or 'png'

def save_tasks_to_excel(df, assembly_name, output_path):
    """
    Save tasks to Excel file with proper formatting.
    
    Args:
        df: DataFrame containing tasks
        assembly_name: Name of the assembly to set as description
        output_path: Path to save the Excel file
    """
    try:
        # Set the description column to the assembly name
        if df is not None and not df.empty:
            # Make column headers match the expected format
            df = df.rename(columns={
                'task_no': 'task_no',
                'type': 'type',
                'eta_sec': 'eta_sec',
                'description': 'description',
                'activity': 'activity',
                'specification': 'specification',
                'attachment': 'attachment'
            })
            
            # Ensure the description is set to assembly name for all rows
            df['description'] = assembly_name
            
            # Create a new Excel workbook with proper formatting
            wb = Workbook()
            ws = wb.active
            
            # Add headers with formatting
            headers = ['task_no', 'type', 'eta_sec', 'description', 'activity', 'specification', 'attachment']
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_idx, value=header)
                # Apply red font color to first 3 columns
                if col_idx <= 3:
                    cell.font = Font(color="FF0000")
            
            # Add data rows
            for row_idx, row in enumerate(df.itertuples(index=False), 2):
                for col_idx, value in enumerate(row, 1):
                    ws.cell(row=row_idx, column=col_idx, value=value)
            
            # Save to Excel
            wb.save(output_path)
            return True
        return False
    except Exception as e:
        logger.error(f"Error saving Excel file: {e}")
        return False

def save_images(images, output_dir):
    """
    Save extracted images to directory.
    
    Args:
        images: Dictionary mapping image IDs to image data
        output_dir: Directory to save images
    """
    try:
        os.makedirs(output_dir, exist_ok=True)
        for image_id, image_info in images.items():
            output_path = os.path.join(output_dir, f"{image_id}.{image_info['extension']}")
            with open(output_path, 'wb') as f:
                f.write(image_info['data'])
        return True
    except Exception as e:
        logger.error(f"Error saving images: {e}")
        return False

def create_zip_package(excel_path, images_dir, output_path, assembly_name):
    """
    Create a ZIP package containing Excel file and images.
    
    Args:
        excel_path: Path to Excel file
        images_dir: Directory containing images
        output_path: Path to save the ZIP file
        assembly_name: Name of the assembly to use for file naming
    """
    try:
        with zipfile.ZipFile(output_path, 'w') as zipf:
            # Add Excel file
            zipf.write(excel_path, arcname=os.path.basename(excel_path))
            
            # Add images
            for root, _, files in os.walk(images_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.join('images', file)
                    zipf.write(file_path, arcname=arcname)
        return True
    except Exception as e:
        logger.error(f"Error creating ZIP package: {e}")
        return False

# Main function to process document
def process_document(input_file, assembly_id, assembly_name, figure_start, figure_end):
    """
    Process a Word document to extract tasks and images.
    
    Args:
        input_file: Path to input Word document
        assembly_id: Assembly sequence ID
        assembly_name: Name of the assembly (description)
        figure_start: Starting figure reference number
        figure_end: Ending figure reference number
    
    Returns:
        Dictionary with processing results
    """
    try:
        # Create temp directory
        temp_dir = os.path.join(os.path.dirname(input_file), "temp_processing")
        os.makedirs(temp_dir, exist_ok=True)
        
        logger.info(f"Processing document with assembly ID: {assembly_id}, name: {assembly_name}")
        logger.info(f"Figure range: {figure_start} to {figure_end}")
        
        # Extract tasks and images
        tasks_df, images = extract_tasks_from_word(
            input_file, 
            assembly_id, 
            assembly_name, 
            figure_start, 
            figure_end
        )
        
        if tasks_df is None or tasks_df.empty:
            return {'success': False, 'message': 'No tasks could be extracted from the document'}
        
        # Save tasks to Excel
        excel_path = os.path.join(temp_dir, f"{assembly_name}.xlsx")
        save_tasks_to_excel(tasks_df, assembly_name, excel_path)
        
        # Save images
        images_dir = os.path.join(temp_dir, "images")
        save_images(images, images_dir)
        
        # Create ZIP package
        zip_path = os.path.join(os.path.dirname(input_file), f"{assembly_name}_extracted_data.zip")
        create_zip_package(excel_path, images_dir, zip_path, assembly_name)
        
        # Return results
        return {
            'success': True,
            'message': f'Successfully processed document. Extracted {len(tasks_df)} tasks and {len(images)} images.',
            'tasks': tasks_df.to_dict('records'),
            'images_count': len(images),
            'excel_path': excel_path,
            'zip_path': zip_path
        }
    except Exception as e:
        logger.error(f"Error in process_document: {str(e)}")
        return {'success': False, 'message': f'Error: {str(e)}'}

# Command line interface
if __name__ == "__main__":
    if len(sys.argv) < 6:
        print("Usage: python script.py input_file assembly_id assembly_name figure_start figure_end")
        sys.exit(1)
    
    input_file = sys.argv[1]
    assembly_id = sys.argv[2]
    assembly_name = sys.argv[3]
    figure_start = int(sys.argv[4])
    figure_end = int(sys.argv[5])
    
    result = process_document(input_file, assembly_id, assembly_name, figure_start, figure_end)
    # Print as JSON for the Node.js wrapper to parse
    print(json.dumps(result))
