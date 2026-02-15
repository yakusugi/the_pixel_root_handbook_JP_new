#!/usr/bin/env python3
"""
Compress images in a .docx file to reduce file size
"""
import os
import zipfile
import shutil
from PIL import Image
from io import BytesIO
import tempfile

def compress_docx_images(input_docx, output_docx, quality=60, max_dimension=1920):
    """
    Compress images in a .docx file
    
    Args:
        input_docx: Path to input .docx file
        output_docx: Path to output .docx file
        quality: JPEG quality (1-100, lower = smaller file)
        max_dimension: Maximum width/height in pixels
    """
    # Create a temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        extract_dir = os.path.join(temp_dir, 'docx_contents')
        
        # Extract the .docx (it's actually a ZIP file)
        print(f"Extracting {input_docx}...")
        with zipfile.ZipFile(input_docx, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
        
        # Find all images in word/media/
        media_dir = os.path.join(extract_dir, 'word', 'media')
        
        if not os.path.exists(media_dir):
            print("No images found in document")
            shutil.copy2(input_docx, output_docx)
            return
        
        # Process each image
        total_original_size = 0
        total_compressed_size = 0
        
        for filename in os.listdir(media_dir):
            filepath = os.path.join(media_dir, filename)
            
            # Skip non-image files
            if not filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff')):
                continue
            
            original_size = os.path.getsize(filepath)
            total_original_size += original_size
            
            try:
                # Open and compress image
                with Image.open(filepath) as img:
                    # Convert to RGB if necessary (for PNG with transparency, etc.)
                    if img.mode in ('RGBA', 'LA', 'P'):
                        # Create white background
                        background = Image.new('RGB', img.size, (255, 255, 255))
                        if img.mode == 'P':
                            img = img.convert('RGBA')
                        background.paste(img, mask=img.split()[-1] if img.mode in ('RGBA', 'LA') else None)
                        img = background
                    elif img.mode != 'RGB':
                        img = img.convert('RGB')
                    
                    # Resize if too large
                    if max(img.size) > max_dimension:
                        ratio = max_dimension / max(img.size)
                        new_size = tuple(int(dim * ratio) for dim in img.size)
                        img = img.resize(new_size, Image.Resampling.LANCZOS)
                        print(f"  Resized {filename} to {new_size}")
                    
                    # Save as JPEG with compression
                    # Change extension to .jpg
                    base_name = os.path.splitext(filename)[0]
                    new_filename = base_name + '.jpg'
                    new_filepath = os.path.join(media_dir, new_filename)
                    
                    # If we're renaming, remove the old file
                    if new_filename != filename:
                        os.remove(filepath)
                        # Update references in document.xml files
                        update_image_references(extract_dir, filename, new_filename)
                    
                    img.save(new_filepath, 'JPEG', quality=quality, optimize=True)
                    
                    compressed_size = os.path.getsize(new_filepath)
                    total_compressed_size += compressed_size
                    
                    reduction = (1 - compressed_size / original_size) * 100
                    print(f"  {filename}: {original_size/1024:.1f}KB -> {compressed_size/1024:.1f}KB ({reduction:.1f}% reduction)")
                    
            except Exception as e:
                print(f"  Error processing {filename}: {e}")
                total_compressed_size += original_size
        
        print(f"\nTotal: {total_original_size/1024/1024:.2f}MB -> {total_compressed_size/1024/1024:.2f}MB")
        
        # Repackage as .docx
        print(f"\nCreating {output_docx}...")
        with zipfile.ZipFile(output_docx, 'w', zipfile.ZIP_DEFLATED) as docx:
            for root, dirs, files in os.walk(extract_dir):
                for file in files:
                    filepath = os.path.join(root, file)
                    arcname = os.path.relpath(filepath, extract_dir)
                    docx.write(filepath, arcname)
        
        final_size = os.path.getsize(output_docx)
        print(f"Final .docx size: {final_size/1024/1024:.2f}MB")

def update_image_references(extract_dir, old_name, new_name):
    """Update image references in document XML files"""
    import re
    
    # Files that might contain image references
    xml_files = [
        'word/document.xml',
        'word/_rels/document.xml.rels',
    ]
    
    for xml_file in xml_files:
        xml_path = os.path.join(extract_dir, xml_file)
        if not os.path.exists(xml_path):
            continue
        
        with open(xml_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Replace references
        content = content.replace(old_name, new_name)
        
        with open(xml_path, 'w', encoding='utf-8') as f:
            f.write(content)

if __name__ == '__main__':
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python3 compress_docx_images.py input.docx [output.docx] [quality] [max_dimension]")
        print("  quality: JPEG quality 1-100 (default: 60)")
        print("  max_dimension: Maximum width/height in pixels (default: 1920)")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else input_file.replace('.docx', '_compressed.docx')
    quality = int(sys.argv[3]) if len(sys.argv) > 3 else 60
    max_dim = int(sys.argv[4]) if len(sys.argv) > 4 else 1920
    
    print(f"Input: {input_file}")
    print(f"Output: {output_file}")
    print(f"Quality: {quality}")
    print(f"Max dimension: {max_dim}px\n")
    
    compress_docx_images(input_file, output_file, quality, max_dim)
    print("\nDone!")
