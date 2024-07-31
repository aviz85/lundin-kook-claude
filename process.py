import os
import json
import requests
import logging
import argparse
from dotenv import load_dotenv
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Load environment variables
load_dotenv()

# Choose model: 'claude-3-opus-20240229' or 'claude-3-5-sonnet-20240620'
MODEL = 'claude-3-5-sonnet-20240620'

# API endpoint
API_URL = 'https://api.anthropic.com/v1/messages'

# Headers
headers = {
    'Content-Type': 'application/json',
    'x-api-key': os.getenv('ANTHROPIC_API_KEY'),
    'anthropic-version': '2023-06-01'
}

# Set up logging
logging.basicConfig(filename='api_usage.log', level=logging.INFO, 
                    format='%(asctime)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

# Initialize usage counters
total_input_tokens = 0
total_output_tokens = 0
model_usage = {}
errors = []

def read_file(filename):
    with open(filename, 'r', encoding='utf-8') as file:
        return file.read()

def write_file(filename, content):
    with open(filename, 'w', encoding='utf-8') as file:
        file.write(content)

def call_claude_api(system_prompt, user_message):
    global total_input_tokens, total_output_tokens, model_usage
    
    data = {
        'model': MODEL,
        'messages': [
            {'role': 'user', 'content': user_message}
        ],
        'system': system_prompt,
        'max_tokens': 4096
    }
    
    try:
        response = requests.post(API_URL, json=data, headers=headers)
        response.raise_for_status()
        response_data = response.json()
        
        # Log usage
        model = response_data['model']
        input_tokens = response_data['usage']['input_tokens']
        output_tokens = response_data['usage']['output_tokens']
        total_input_tokens += input_tokens
        total_output_tokens += output_tokens
        
        # Update model-specific usage
        if model not in model_usage:
            model_usage[model] = {'input_tokens': 0, 'output_tokens': 0}
        model_usage[model]['input_tokens'] += input_tokens
        model_usage[model]['output_tokens'] += output_tokens
        
        logging.info(f"API call - Model: {model}, Input tokens: {input_tokens}, Output tokens: {output_tokens}")
        
        return response_data['content'][0]['text']
    except Exception as e:
        error_msg = f"Error in API call: {str(e)}"
        errors.append(error_msg)
        logging.error(error_msg)
        return None

def set_rtl(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    biDi = OxmlElement('w:bidi')
    pPr.insert_element_before(biDi, 'w:jc')

def compile_to_docx(results_dir, output_file):
    doc = Document()
    
    # Set RTL direction for the entire document
    section = doc.sections[0]
    section.page_width, section.page_height = section.page_height, section.page_width
    
    # Set default font
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(12)
    
    for filename in sorted(os.listdir(results_dir)):
        if filename.endswith('.json'):
            try:
                with open(os.path.join(results_dir, filename), 'r', encoding='utf-8') as file:
                    data = json.load(file)
                
                # Add letter
                para = doc.add_paragraph()
                set_rtl(para)
                para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                run = para.add_run(data.get('letter', ''))
                run.bold = True
                doc.add_paragraph()
                
                # Add original text
                para = doc.add_paragraph()
                set_rtl(para)
                para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                para.add_run(data['original_text'])
                doc.add_paragraph()
                
                # Add difficult words explanations
                difficult_words = data['difficult_words']
                if difficult_words:
                    para = doc.add_paragraph()
                    set_rtl(para)
                    para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                    explanations = '; '.join([f"{item['word']} – {item['explanation']}" for item in difficult_words])
                    para.add_run(explanations)
                    doc.add_paragraph()
                
                # Add detailed interpretation
                detailed_interpretation = data['detailed_interpretation']
                para = doc.add_paragraph()
                set_rtl(para)
                para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                for part in detailed_interpretation:
                    run = para.add_run(part['quote'])
                    run.bold = True
                    para.add_run(f" - {part['explanation']} ")
                
                # Add extra paragraph for spacing between sections
                doc.add_paragraph()
            except Exception as e:
                error_msg = f"Error processing file {filename}: {str(e)}"
                errors.append(error_msg)
                logging.error(error_msg)
    
    doc.save(output_file)
    print(f"Compiled document saved as {output_file}")

def main():
    parser = argparse.ArgumentParser(description="Process text files and compile into DOCX.")
    parser.add_argument("--skip-processing", action="store_true", help="Skip processing and only compile existing JSON files")
    args = parser.parse_args()

    if not args.skip_processing:
        # Read prompt and examples
        prompt = read_file('prompt.txt')
        examples = read_file('examples.txt')
        
        # Example of correct interpretation structure
        example_structure = {
            "letter": "א",
            "original_text": "התכונה של יראת שמים, מצד עצמה, לית לה מגרמה כלום, ואי אפשר לה להיות מתחשבת בין הכשרונות ומעלות הנפש של האדם.",
            "difficult_words": [
                {"word": "לית לה מגרמה כלום", "explanation": "אין לה מעצמה כלום"},
                {"word": "מתחשבת", "explanation": "נחשבת, נספרת"}
            ],
            "detailed_interpretation": [
                {
                    "quote": "התכונה של יראת שמים, מצד עצמה, לית לה מגרמה כלום",
                    "explanation": "השאיפה הדתית (\"יראת שמים\") איננה תוכן העומד בפני עצמו. הרב קוק מסביר כי יראת שמים, כשלעצמה, אינה בעלת ערך עצמאי."
                },
                {
                    "quote": "ואי אפשר לה להיות מתחשבת בין הכשרונות ומעלות הנפש של האדם",
                    "explanation": "יראת שמים איננה נספרת בין שאר כוחות הנפש. היא אינה יכולה להיחשב כאחת מהתכונות או היכולות של האדם."
                }
            ]
        }
        
        # Combine prompt, examples, and JSON structure example
        system_prompt = f"""{prompt}

    פסקאות לדוגמא:

    {examples}

    Please provide your interpretation in JSON format. Here's an example of the correct structure:

    {json.dumps(example_structure, ensure_ascii=False, indent=2)}

    Make sure to follow this structure in your response, using JSON mode."""
        
        # Log the system prompt
        print("System Prompt:")
        print("=" * 50)
        print(system_prompt[:500] + "..." if len(system_prompt) > 500 else system_prompt)
        print("=" * 50)
        
        # Create results directory if it doesn't exist
        os.makedirs('results', exist_ok=True)
        
        # Process all txt files in the sources directory
        for filename in os.listdir('sources'):
            if filename.endswith('.txt'):
                print(f"\nProcessing {filename}...")
                
                # Read the paragraph
                paragraph = read_file(os.path.join('sources', filename))
                
                # Log the current paragraph
                print("Current Paragraph:")
                print("-" * 50)
                print(paragraph[:500] + "..." if len(paragraph) > 500 else paragraph)
                print("-" * 50)
                
                # Call Claude API
                print("Calling Claude API...")
                response = call_claude_api(system_prompt, paragraph)
                
                if response:
                    # Determine output filename
                    output_filename = f"results/{os.path.splitext(filename)[0]}_{MODEL}.json"
                    
                    # Write the response to a file
                    write_file(output_filename, response)
                    
                    print(f"Result saved to {output_filename}")
                    
                    # Log a snippet of the response
                    print("Response snippet:")
                    print("~" * 50)
                    print(response[:500] + "..." if len(response) > 500 else response)
                    print("~" * 50)
                else:
                    print(f"Failed to process {filename}")
    
    # Compile all JSON files into a single DOCX
    compile_to_docx('results', 'compiled_interpretations.docx')
    
    # Log total usage
    logging.info("Total usage:")
    for model, usage in model_usage.items():
        logging.info(f"Model: {model} - Input tokens: {usage['input_tokens']}, Output tokens: {usage['output_tokens']}")
    logging.info(f"Overall - Input tokens: {total_input_tokens}, Output tokens: {total_output_tokens}")
    
    print("Total usage:")
    for model, usage in model_usage.items():
        print(f"Model: {model} - Input tokens: {usage['input_tokens']}, Output tokens: {usage['output_tokens']}")
    print(f"Overall - Input tokens: {total_input_tokens}, Output tokens: {total_output_tokens}")
    
    if errors:
        print("\nErrors encountered during processing:")
        for error in errors:
            print(error)

if __name__ == "__main__":
    main()