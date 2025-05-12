
import os
import re
import sys
import docx
import pandas as pd
from docx.document import Document as DocxDocument
from PIL import Image
from io import BytesIO
import base64
import zipfile

def extract_tasks_from_word(docx_path, assembly_id, figure_start_range, figure_end_range):
    """
    Extract tasks from a Word document that contains a table.
    
    Args:
        docx_path: Path to the Word document
        assembly_id: Assembly sequence ID to prefix task numbers
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
        image_counter = 1
        
        # Find tables in the document
        if len(doc.tables) == 0:
            print("No tables found in the document")
            return None, {}
        
        # Assume the first table contains the tasks
        table = doc.tables[0]
        
        # Process each row of the table (skipping header if exists)
        for i, row in enumerate(table.rows):
            # Skip if it's likely a header row
            if i == 0 and any(cell.text.lower() in ['sl no', 'sl.no', 'sl. no', 'serial no', 'task no'] for cell in row.cells):
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
                    print(f"Warning: Could not parse sl_no from '{sl_no_cell}', skipping row")
                    continue
                
                # Generate task number in the format assembly_id-0-XXX
                task_number = f"{assembly_id}-0-{str(sl_no).zfill(3)}"
                
                # Extract job details text
                job_details = ""
                for paragraph in job_details_cell.paragraphs:
                    job_details += paragraph.text + "\n"
                job_details = job_details.strip()
                
                # Extract image references
                image_references = []
                figure_pattern = re.compile(r"Figure\s+(\d+)", re.IGNORECASE)
                matches = figure_pattern.findall(job_details)
                
                for match in matches:
                    fig_num = int(match)
                    if figure_start_range <= fig_num <= figure_end_range:
                        # Calculate image ID based on the figure reference
                        image_id = f"{assembly_id}-0-{str(image_counter).zfill(3)}"
                        image_counter += 1
                        image_references.append(image_id)
                
                # Create task entry
                task = {
                    'task_no': task_number,
                    'type': 'Operation',
                    'eta_sec': '',
                    'description': '',  # Will be filled later with assembly name
                    'activity': job_details,
                    'specification': '',
                    'attachment': ', '.join(image_references)
                }
                
                tasks.append(task)
            except Exception as e:
                print(f"Error processing row {i}: {e}")
                continue
        
        # Extract images
        image_counter = 1
        for rel in doc.part.rels.values():
            if "image" in rel.target_ref:
                try:
                    image_data = rel.target_part.blob
                    # Determine file extension based on image data
                    import imghdr
                    img_ext = imghdr.what(None, h=image_data) or 'png'
                    
                    # Generate image ID
                    image_id = f"{assembly_id}-0-{str(image_counter).zfill(3)}"
                    image_counter += 1
                    
                    # Store image data
                    images[image_id] = {
                        'data': image_data,
                        'extension': img_ext
                    }
                except Exception as e:
                    print(f"Error extracting image: {e}")
        
        # Create DataFrame
        if tasks:
            df = pd.DataFrame(tasks)
            return df, images
        else:
            print("No tasks extracted from the document")
            return None, {}
            
    except Exception as e:
        print(f"Error processing document: {e}")
        return None, {}

def save_tasks_to_excel(df, assembly_name, output_path):
    """
    Save tasks to Excel file.
    
    Args:
        df: DataFrame containing tasks
        assembly_name: Name of the assembly to set as description
        output_path: Path to save the Excel file
    """
    try:
        # Set the description column to the assembly name
        if df is not None and not df.empty:
            df['description'] = assembly_name
            df.to_excel(output_path, index=False)
            return True
        return False
    except Exception as e:
        print(f"Error saving Excel file: {e}")
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
        print(f"Error saving images: {e}")
        return False

def create_zip_package(excel_path, images_dir, output_path):
    """
    Create a ZIP package containing Excel file and images.
    
    Args:
        excel_path: Path to Excel file
        images_dir: Directory containing images
        output_path: Path to save the ZIP file
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
        print(f"Error creating ZIP package: {e}")
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
        
        # Extract tasks and images
        tasks_df, images = extract_tasks_from_word(
            input_file, 
            assembly_id, 
            figure_start, 
            figure_end
        )
        
        if tasks_df is None or tasks_df.empty:
            return {'success': False, 'message': 'No tasks could be extracted from the document'}
        
        # Save tasks to Excel
        excel_path = os.path.join(temp_dir, f"Tasks_{assembly_name}.xlsx")
        save_tasks_to_excel(tasks_df, assembly_name, excel_path)
        
        # Save images
        images_dir = os.path.join(temp_dir, "images")
        save_images(images, images_dir)
        
        # Create ZIP package
        zip_path = os.path.join(os.path.dirname(input_file), f"{assembly_name}_Package.zip")
        create_zip_package(excel_path, images_dir, zip_path)
        
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
    print(result)
