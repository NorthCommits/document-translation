import json
from collections import Counter

def analyze_json_structure(json_file_path, output_file="summary_report.txt"):
    """
    Analyze extracted PowerPoint JSON and generate a concise summary report
    """
    
    with open(json_file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    summary = []
    summary.append("=" * 80)
    summary.append("POWERPOINT EXTRACTION SUMMARY REPORT")
    summary.append("=" * 80)
    summary.append(f"\nPresentation: {data.get('presentation_name', 'Unknown')}")
    summary.append(f"Total Slides: {data.get('total_slides', 0)}")
    summary.append("\n")
    
    # Statistics
    total_elements = 0
    element_types = Counter()
    total_paragraphs = 0
    total_runs = 0
    bullet_types = Counter()
    font_names = Counter()
    font_sizes = Counter()
    text_sample_count = 0
    
    # Analyze slides
    for slide_idx, slide in enumerate(data.get('slides', []), 1):
        total_elements += len(slide.get('elements', []))
        
        for element in slide.get('elements', []):
            element_types[element.get('element_type', 'Unknown')] += 1
            
            # Analyze paragraphs
            for para in element.get('paragraphs', []):
                total_paragraphs += 1
                
                # Bullet info
                bullet_format = para.get('paragraph_formatting', {}).get('bullet_format', {})
                if bullet_format.get('is_bulleted'):
                    bullet_types[bullet_format.get('bullet_type', 'bullet')] += 1
                
                # Analyze runs
                for run in para.get('runs', []):
                    total_runs += 1
                    
                    if run.get('font_name'):
                        font_names[run['font_name']] += 1
                    if run.get('font_size'):
                        font_sizes[run['font_size']] += 1
    
    # Write summary
    summary.append("-" * 80)
    summary.append("OVERALL STATISTICS")
    summary.append("-" * 80)
    summary.append(f"Total Elements: {total_elements}")
    summary.append(f"Total Paragraphs: {total_paragraphs}")
    summary.append(f"Total Text Runs: {total_runs}")
    summary.append(f"Total Links: {sum(len(s.get('links', [])) for s in data.get('slides', []))}")
    summary.append(f"Slides with Speaker Notes: {sum(1 for s in data.get('slides', []) if s.get('speaker_notes'))}")
    summary.append(f"SmartArt Objects: {sum(len(s.get('smartart', [])) for s in data.get('slides', []))}")
    
    summary.append("\n" + "-" * 80)
    summary.append("ELEMENT TYPES BREAKDOWN")
    summary.append("-" * 80)
    for elem_type, count in element_types.most_common():
        summary.append(f"  {elem_type}: {count}")
    
    summary.append("\n" + "-" * 80)
    summary.append("BULLET/NUMBERING STATISTICS")
    summary.append("-" * 80)
    summary.append(f"Total Bulleted/Numbered Paragraphs: {sum(bullet_types.values())}")
    for bullet_type, count in bullet_types.most_common():
        summary.append(f"  {bullet_type}: {count}")
    
    summary.append("\n" + "-" * 80)
    summary.append("TOP 10 FONTS USED")
    summary.append("-" * 80)
    for font_name, count in font_names.most_common(10):
        summary.append(f"  {font_name}: {count} times")
    
    summary.append("\n" + "-" * 80)
    summary.append("FONT SIZES USED")
    summary.append("-" * 80)
    for font_size, count in sorted(font_sizes.items()):
        summary.append(f"  {font_size}pt: {count} times")
    
    # Sample elements from each slide
    summary.append("\n" + "=" * 80)
    summary.append("SLIDE-BY-SLIDE SAMPLES (First 3 Slides)")
    summary.append("=" * 80)
    
    for slide_idx, slide in enumerate(data.get('slides', [])[:3], 1):
        summary.append(f"\n{'=' * 80}")
        summary.append(f"SLIDE {slide_idx}")
        summary.append(f"{'=' * 80}")
        summary.append(f"Elements: {len(slide.get('elements', []))}")
        summary.append(f"Links: {len(slide.get('links', []))}")
        summary.append(f"Has Speaker Notes: {'Yes' if slide.get('speaker_notes') else 'No'}")
        
        # Show first 2 elements in detail
        for elem_idx, element in enumerate(slide.get('elements', [])[:2], 1):
            summary.append(f"\n  Element {elem_idx}: {element.get('element_type', 'Unknown')}")
            summary.append(f"    Shape ID: {element.get('shape_id')}")
            summary.append(f"    Shape Name: {element.get('shape_name')}")
            
            # Show first paragraph with formatting
            if element.get('paragraphs'):
                first_para = element['paragraphs'][0]
                para_format = first_para.get('paragraph_formatting', {})
                
                summary.append(f"    Paragraph Format:")
                summary.append(f"      Level: {para_format.get('level')}")
                summary.append(f"      Alignment: {para_format.get('alignment')}")
                
                bullet_info = para_format.get('bullet_format', {})
                if bullet_info.get('is_bulleted'):
                    summary.append(f"      Bullet Type: {bullet_info.get('bullet_type')}")
                    summary.append(f"      Bullet Char: {bullet_info.get('bullet_char')}")
                    summary.append(f"      Bullet Color: {bullet_info.get('bullet_color')}")
                
                # Show first run
                if first_para.get('runs'):
                    first_run = first_para['runs'][0]
                    summary.append(f"    First Run Format:")
                    summary.append(f"      Font: {first_run.get('font_name')}")
                    summary.append(f"      Size: {first_run.get('font_size')}pt")
                    summary.append(f"      Bold: {first_run.get('bold')}")
                    summary.append(f"      Italic: {first_run.get('italic')}")
                    summary.append(f"      Color: {first_run.get('color')}")
                    summary.append(f"      Text Preview: {first_run.get('text', '')[:100]}...")
            
            # Show table structure if it's a table
            if element.get('element_type') == 'Table':
                summary.append(f"    Table: {element.get('rows')} rows x {element.get('columns')} columns")
                summary.append(f"    Cells with content: {len(element.get('cells', []))}")
    
    # Issues/Missing Data Report
    summary.append("\n" + "=" * 80)
    summary.append("DATA QUALITY CHECK")
    summary.append("=" * 80)
    
    issues = []
    
    # Check for missing SmartArt
    smartart_count = sum(len(s.get('smartart', [])) for s in data.get('slides', []))
    if smartart_count == 0:
        issues.append("⚠ No SmartArt objects found (may not exist or extraction failed)")
    
    # Check for theme colors
    theme_color_count = 0
    rgb_color_count = 0
    for slide in data.get('slides', []):
        for element in slide.get('elements', []):
            for para in element.get('paragraphs', []):
                for run in para.get('runs', []):
                    color = run.get('color')
                    if color:
                        if isinstance(color, dict):
                            if 'theme_color' in color:
                                theme_color_count += 1
                            if 'rgb' in color:
                                rgb_color_count += 1
    
    summary.append(f"\nColor Usage:")
    summary.append(f"  RGB Colors: {rgb_color_count}")
    summary.append(f"  Theme Colors: {theme_color_count}")
    
    if issues:
        summary.append("\nPotential Issues:")
        for issue in issues:
            summary.append(f"  {issue}")
    else:
        summary.append("\n✓ No major issues detected")
    
    # Write to file
    summary_text = "\n".join(summary)
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(summary_text)
    
    print(f"Summary report generated: {output_file}")
    print(f"\nReport is {len(summary_text)} characters")
    print("\n" + "=" * 80)
    print("QUICK STATS:")
    print("=" * 80)
    print(f"Slides: {data.get('total_slides', 0)}")
    print(f"Elements: {total_elements}")
    print(f"Paragraphs: {total_paragraphs}")
    print(f"Text Runs: {total_runs}")
    
    return summary_text


def extract_sample_slide(json_file_path, slide_number=1, output_file="sample_slide.json"):
    """
    Extract a single slide as a sample for review
    """
    with open(json_file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    if slide_number <= len(data.get('slides', [])):
        sample = {
            "presentation_name": data.get('presentation_name'),
            "slide": data['slides'][slide_number - 1]
        }
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(sample, f, indent=2, ensure_ascii=False)
        
        print(f"Sample slide {slide_number} extracted to: {output_file}")
    else:
        print(f"Slide {slide_number} not found. Total slides: {len(data.get('slides', []))}")


if __name__ == "__main__":
    # Replace with your JSON file path
    json_file = "extracted_content_enhanced.json"
    
    # Generate summary report
    summary = analyze_json_structure(json_file, "summary_report.txt")
    
    # Extract first slide as sample
    extract_sample_slide(json_file, slide_number=1, output_file="sample_slide_1.json")
    
    print("\n✓ Done! You can now share 'summary_report.txt' for review")