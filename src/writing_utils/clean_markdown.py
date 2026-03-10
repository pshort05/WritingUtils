#!/usr/bin/env python3
import argparse
import sys
import re

def is_markdown_structure(line):
    """
    Identifies if a line starts with Markdown syntax that should not be indented.
    """
    stripped = line.lstrip()
    if not stripped:
        return False
    
    # Headers: #, ##, etc.
    if stripped.startswith('#'):
        return True
    
    # Lists: *, -, +, or 1., 2., etc.
    if re.match(r'^([\*\-\+]|\d+\.)\s', stripped):
        return True
    
    # Blockquotes: >
    if stripped.startswith('>'):
        return True
    
    # Code blocks: ``` or ~~~
    if stripped.startswith('```') or stripped.startswith('~~~'):
        return True
    
    # Horizontal rules: ---, ***, ___
    if re.match(r'^(\-{3,}|\*{3,}|_{3,})\s*$', stripped):
        return True
    
    # Images and Links
    if stripped.startswith('![') or stripped.startswith('['):
        return True
        
    # HTML tags (basic check)
    if stripped.startswith('<') and '>' in stripped:
        return True
        
    # Novel scene metadata (Bold or Italic only lines)
    if re.match(r'^(\*\*|\*|_).+(\*\*|\*|_)\s*$', stripped):
        return True

    return False

def clean_markdown(content):
    lines = content.splitlines()
    processed_lines = []
    
    # Identify blank lines
    line_data = [{'text': l, 'is_blank': not l.strip()} for l in lines]
    
    final_output = []
    i = 0
    while i < len(line_data):
        line_obj = line_data[i]
        
        if line_obj['is_blank']:
            # Count consecutive blank lines
            blank_group = []
            j = i
            while j < len(line_data) and line_data[j]['is_blank']:
                blank_group.append(line_data[j]['text'])
                j += 1
            
            blank_count = len(blank_group)
            
            # Contextual check for previous and next lines
            prev_content_line = ""
            for k in range(len(final_output) - 1, -1, -1):
                if final_output[k].strip():
                    prev_content_line = final_output[k]
                    break
            
            next_content_line = ""
            for k in range(j, len(line_data)):
                if not line_data[k]['is_blank']:
                    next_content_line = line_data[k]['text']
                    break

            # Identify if adjacent lines are structural
            prev_structural = is_markdown_structure(prev_content_line) if prev_content_line else False
            next_structural = is_markdown_structure(next_content_line) if next_content_line else False
            
            # Rule 1: Keep blank lines around structural elements or if multiple
            if prev_structural or next_structural or blank_count > 1:
                final_output.extend(blank_group)
            else:
                # Single blank line between regular paragraphs: Remove it
                pass
            
            i = j  # Move to the next non-blank line
        else:
            # Content line
            text = line_obj['text']
            
            # Rule 2: Indents each new paragraph
            # A "new paragraph" is the first line of the file or a line following a blank line in the source.
            is_new_para = False
            if i == 0:
                is_new_para = True
            elif line_data[i-1]['is_blank']:
                is_new_para = True
            
            # Rule 3: Do not touch Markdown syntax
            if is_new_para and not is_markdown_structure(text):
                # Apply indentation (4 spaces)
                text = "    " + text.lstrip()
            
            final_output.append(text)
            i += 1
            
    return "\n".join(final_output)

def main():
    parser = argparse.ArgumentParser(description="Clean up Markdown files with specific indentation and blank line rules.")
    parser.add_argument("-i", "--input", required=True, help="Input Markdown file")
    parser.add_argument("-o", "--output", required=True, help="Output Markdown file")
    
    args = parser.parse_args()
    
    try:
        with open(args.input, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as e:
        print(f"Error reading input file: {e}", file=sys.stderr)
        sys.exit(1)
        
    cleaned_content = clean_markdown(content)
    
    try:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(cleaned_content)
            if not cleaned_content.endswith('\n'):
                f.write('\n')
    except Exception as e:
        print(f"Error writing output file: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
