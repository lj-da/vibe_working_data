#!/usr/bin/env python3
"""
Script to add "_gt" suffix to filename in expected.path of JSON files
Can process single file or entire directories
"""
import json
import os
import sys
from pathlib import Path

def add_gt_suffix_to_json(json_path):
    """Add '_gt' suffix to filename in expected.path"""
    json_path = Path(json_path)
    
    if not json_path.exists():
        return False, f"File not found: {json_path}"
    
    try:
        # Read JSON file
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Check if expected.path exists
        if 'evaluator' not in data or 'expected' not in data['evaluator']:
            return False, f"No 'evaluator.expected' found"
        
        expected = data['evaluator']['expected']
        if 'path' not in expected:
            return False, f"No 'path' in expected section"
        
        original_path = expected['path']
        
        # Extract filename from path
        # Handle both URL and local path
        if '/' in original_path:
            path_parts = original_path.rsplit('/', 1)
            if len(path_parts) == 2:
                directory = path_parts[0]
                filename = path_parts[1]
            else:
                directory = ''
                filename = path_parts[0]
        else:
            directory = ''
            filename = original_path
        
        # Skip if already has _gt suffix
        if '_gt.' in filename or filename.endswith('_gt'):
            return True, "Already has _gt suffix"
        
        # Add _gt before extension
        if '.' in filename:
            name, ext = filename.rsplit('.', 1)
            new_filename = f"{name}_gt.{ext}"
        else:
            new_filename = f"{filename}_gt"
        
        # Reconstruct path
        if directory:
            new_path = f"{directory}/{new_filename}"
        else:
            new_path = new_filename
        
        # Also update dest if it exists
        if 'dest' in expected:
            original_dest = expected['dest']
            if '_gt.' not in original_dest and not original_dest.endswith('_gt'):
                if '.' in original_dest:
                    dest_name, dest_ext = original_dest.rsplit('.', 1)
                    new_dest = f"{dest_name}_gt.{dest_ext}"
                else:
                    new_dest = f"{original_dest}_gt"
                expected['dest'] = new_dest
        
        # Update path
        expected['path'] = new_path
        
        # Write back to file
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        
        return True, f"Updated: {filename} -> {new_filename}"
    
    except Exception as e:
        return False, f"Error: {str(e)}"

def process_directory(directory_path):
    """Process all JSON files in a directory"""
    directory = Path(directory_path)
    if not directory.exists() or not directory.is_dir():
        print(f"Error: Directory not found: {directory}")
        return 0, 0
    
    json_files = list(directory.glob('*.json'))
    if not json_files:
        print(f"No JSON files found in {directory}")
        return 0, 0
    
    success_count = 0
    error_count = 0
    
    print(f"\nProcessing {len(json_files)} files in {directory}...")
    
    for json_file in sorted(json_files):
        success, message = add_gt_suffix_to_json(json_file)
        if success:
            success_count += 1
            print(f"  ✓ {json_file.name}: {message}")
        else:
            error_count += 1
            print(f"  ✗ {json_file.name}: {message}")
    
    return success_count, error_count

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python add_gt_suffix.py <json_file_path_or_directory>")
        print("       python add_gt_suffix.py --batch <dir1> <dir2> ...")
        sys.exit(1)
    
    if sys.argv[1] == '--batch':
        # Batch mode: process multiple directories
        directories = sys.argv[2:]
        if not directories:
            directories = [
                'evaluation_examples/examples/excel_visualization',
                'evaluation_examples/examples/excel_100',
                'evaluation_examples/examples/excel_template'
            ]
        
        total_success = 0
        total_errors = 0
        
        for directory in directories:
            success, errors = process_directory(directory)
            total_success += success
            total_errors += errors
        
        print(f"\n{'='*60}")
        print(f"Summary: {total_success} files updated, {total_errors} errors")
        print(f"{'='*60}")
        sys.exit(0 if total_errors == 0 else 1)
    else:
        # Single file or directory mode
        target = Path(sys.argv[1])
        
        if target.is_file():
            success, message = add_gt_suffix_to_json(target)
            if success:
                print(f"✓ {message}")
                sys.exit(0)
            else:
                print(f"✗ {message}")
                sys.exit(1)
        elif target.is_dir():
            success_count, error_count = process_directory(target)
            print(f"\nSummary: {success_count} files updated, {error_count} errors")
            sys.exit(0 if error_count == 0 else 1)
        else:
            print(f"Error: {target} is not a valid file or directory")
            sys.exit(1)

