import json
import os
from openai import OpenAI
from typing import Dict, List, Any
from dotenv import load_dotenv
import time
from copy import deepcopy

class PPTTranslator:
    """
    Translates PowerPoint extracted content while preserving 100% of metadata.
    Only translates actual text content, keeping all formatting and structural data intact.
    
    Updated to handle comprehensive extraction with:
    - Slide masters and layouts
    - Fill, line, and shadow properties
    - Placeholder information
    - Background details
    """
    
    def __init__(self, api_key: str = None, target_language: str = "Spanish"):
        """
        Initialize the translator.
        
        Args:
            api_key: OpenAI API key (if None, loads from .env)
            target_language: Target language for translation (default: Spanish)
        """
        # Load environment variables
        load_dotenv()
        
        # Get API key
        self.api_key = api_key or os.getenv('OPENAI_API_KEY')
        if not self.api_key:
            raise ValueError("OPENAI_API_KEY not found. Please set it in .env file or pass it as parameter.")
        
        # Initialize OpenAI client
        self.client = OpenAI(api_key=self.api_key)
        self.target_language = target_language
        self.model = "gpt-4o-mini"
        
        # Statistics
        self.stats = {
            "total_texts_translated": 0,
            "api_calls": 0,
            "total_tokens_used": 0
        }
    
    def translate_batch(self, texts: List[str]) -> List[str]:
        """
        Translate a batch of texts using GPT-4o-mini.
        
        Args:
            texts: List of text strings to translate
            
        Returns:
            List of translated text strings in the same order
        """
        if not texts:
            return []
        
        # Filter out empty texts but remember their positions
        text_map = {}
        non_empty_texts = []
        for idx, text in enumerate(texts):
            if text and text.strip():
                text_map[len(non_empty_texts)] = idx
                non_empty_texts.append(text)
        
        if not non_empty_texts:
            return texts
        
        # Use JSON format for more reliable parsing
        texts_json = []
        for idx, text in enumerate(non_empty_texts):
            texts_json.append({
                "id": idx,
                "text": text
            })
        
        import json as json_module
        batch_json = json_module.dumps(texts_json, ensure_ascii=False)
        
        prompt = f"""Translate the texts in the following JSON array to {self.target_language}.

CRITICAL RULES:
1. Return ONLY a JSON array with the same structure
2. Keep the same "id" values
3. Translate only the "text" field
4. Preserve all line breaks (\\n) and special characters
5. Do not add any explanations or extra content outside the JSON
6. The number of items in output must match the input exactly

Input JSON:
{batch_json}

Output (JSON array only):"""

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": f"You are a professional translator. Return only valid JSON. Translate to {self.target_language}."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3,
                max_tokens=4096
            )
            
            # Update statistics
            self.stats["api_calls"] += 1
            self.stats["total_tokens_used"] += response.usage.total_tokens
            
            # Parse response
            response_text = response.choices[0].message.content.strip()
            
            # Extract JSON from response (handle markdown code blocks and extra text)
            if "```json" in response_text:
                response_text = response_text.split("```json")[1].split("```")[0].strip()
            elif "```" in response_text:
                response_text = response_text.split("```")[1].split("```")[0].strip()
            
            # Try to find JSON array boundaries
            if not response_text.startswith('['):
                start_idx = response_text.find('[')
                if start_idx != -1:
                    response_text = response_text[start_idx:]
            
            if not response_text.endswith(']'):
                end_idx = response_text.rfind(']')
                if end_idx != -1:
                    response_text = response_text[:end_idx + 1]
            
            # Parse JSON
            try:
                translated_json = json_module.loads(response_text)
            except json_module.JSONDecodeError:
                import re
                match = re.search(r'\[.*\]', response_text, re.DOTALL)
                if match:
                    response_text = match.group(0)
                    translated_json = json_module.loads(response_text)
                else:
                    raise
            
            # Verify structure
            if not isinstance(translated_json, list) or len(translated_json) != len(non_empty_texts):
                if isinstance(translated_json, list) and all(isinstance(item, dict) and 'id' in item for item in translated_json):
                    translated_json = sorted(translated_json, key=lambda x: x.get('id', 0))
                    translated_texts = [item.get('text', non_empty_texts[i]) for i, item in enumerate(translated_json[:len(non_empty_texts)])]
                else:
                    return texts
            else:
                translated_json = sorted(translated_json, key=lambda x: x.get('id', 0))
                translated_texts = [item.get('text', '') for item in translated_json]
            
            # Reconstruct full list with empty texts in original positions
            result = texts.copy()
            for new_idx, orig_idx in text_map.items():
                if new_idx < len(translated_texts):
                    result[orig_idx] = translated_texts[new_idx]
            
            self.stats["total_texts_translated"] += len(non_empty_texts)
            return result
            
        except json_module.JSONDecodeError:
            return self.translate_one_by_one(texts)
        except Exception as e:
            print(f"Translation error: {e}")
            return self.translate_one_by_one(texts)
    
    def translate_one_by_one(self, texts: List[str]) -> List[str]:
        """
        Fallback method: translate texts one by one.
        
        Args:
            texts: List of text strings to translate
            
        Returns:
            List of translated text strings
        """
        translated = []
        for text in texts:
            if not text or not text.strip():
                translated.append(text)
                continue
            
            try:
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[
                        {"role": "system", "content": f"You are a professional translator. Translate to {self.target_language}. Return ONLY the translated text, nothing else."},
                        {"role": "user", "content": f"Translate this to {self.target_language}:\n\n{text}"}
                    ],
                    temperature=0.3,
                    max_tokens=2048
                )
                
                self.stats["api_calls"] += 1
                self.stats["total_tokens_used"] += response.usage.total_tokens
                self.stats["total_texts_translated"] += 1
                
                translated_text = response.choices[0].message.content.strip()
                translated.append(translated_text)
                
            except Exception as e:
                print(f"Error translating individual text: {e}")
                translated.append(text)
        
        return translated
    
    def translate_text_runs(self, runs: List[Dict]) -> List[Dict]:
        """
        Translate text runs while preserving all formatting metadata.
        
        Args:
            runs: List of run dictionaries containing text and formatting
            
        Returns:
            List of run dictionaries with translated text
        """
        if not runs:
            return runs
        
        # Extract texts for translation
        texts = [run.get("text", "") for run in runs]
        
        # Translate
        translated_texts = self.translate_batch(texts)
        
        # Create new runs with translated text but original metadata
        translated_runs = []
        for idx, run in enumerate(runs):
            new_run = deepcopy(run)
            new_run["text"] = translated_texts[idx]
            translated_runs.append(new_run)
        
        return translated_runs
    
    def translate_paragraphs(self, paragraphs: List[Dict]) -> List[Dict]:
        """
        Translate paragraphs while preserving all paragraph formatting.
        
        Args:
            paragraphs: List of paragraph dictionaries
            
        Returns:
            List of paragraph dictionaries with translated text
        """
        if not paragraphs:
            return paragraphs
        
        translated_paragraphs = []
        for para in paragraphs:
            new_para = deepcopy(para)
            
            # Translate runs
            if "runs" in new_para:
                new_para["runs"] = self.translate_text_runs(new_para["runs"])
            
            translated_paragraphs.append(new_para)
        
        return translated_paragraphs
    
    def translate_text_element(self, element: Dict) -> Dict:
        """
        Translate a text element (TextBox, AutoShape, etc.) preserving all metadata.
        Now handles new fields: fill, line, shadow, placeholder_info
        
        Args:
            element: Element dictionary
            
        Returns:
            Element dictionary with translated text
        """
        new_element = deepcopy(element)
        
        # Translate paragraphs (text content)
        if "paragraphs" in new_element:
            new_element["paragraphs"] = self.translate_paragraphs(new_element["paragraphs"])
        
        # Update full_text by concatenating translated runs
        if "paragraphs" in new_element:
            all_text = []
            for para in new_element["paragraphs"]:
                para_text = []
                if "runs" in para:
                    for run in para["runs"]:
                        if run.get("text"):
                            para_text.append(run["text"])
                if para_text:
                    all_text.append("".join(para_text))
            new_element["full_text"] = "\n".join(all_text) if all_text else ""
        
        # All other fields (fill, line, shadow, placeholder_info, dimensions) are preserved as-is
        
        return new_element
    
    def translate_table(self, table_data: Dict) -> Dict:
        """
        Translate table cells while preserving table structure.
        
        Args:
            table_data: Table data dictionary
            
        Returns:
            Table data dictionary with translated cells
        """
        new_table = deepcopy(table_data)
        
        if "cells" in new_table:
            translated_cells = []
            for cell in new_table["cells"]:
                # Each cell has paragraphs
                if "paragraphs" in cell:
                    cell["paragraphs"] = self.translate_paragraphs(cell["paragraphs"])
                
                # Update cell text
                if "text" in cell and "paragraphs" in cell:
                    all_text = []
                    for para in cell["paragraphs"]:
                        if "runs" in para:
                            para_text = "".join(run.get("text", "") for run in para["runs"])
                            if para_text:
                                all_text.append(para_text)
                    cell["text"] = "\n".join(all_text) if all_text else ""
                
                translated_cells.append(cell)
            new_table["cells"] = translated_cells
        
        return new_table
    
    def translate_chart(self, chart_data: Dict) -> Dict:
        """
        Translate chart text elements while preserving chart data and structure.
        
        Args:
            chart_data: Chart data dictionary
            
        Returns:
            Chart data dictionary with translated text
        """
        new_chart = deepcopy(chart_data)
        
        # Translate chart title
        if "title" in new_chart and new_chart["title"]:
            translated = self.translate_batch([new_chart["title"]])
            new_chart["title"] = translated[0]
        
        # Translate series names in data_values
        if "data_values" in new_chart and new_chart["data_values"]:
            series_names = [s.get("series_name", "") for s in new_chart["data_values"] if s.get("series_name")]
            if series_names:
                translated_names = self.translate_batch(series_names)
                name_idx = 0
                for series in new_chart["data_values"]:
                    if series.get("series_name"):
                        series["series_name"] = translated_names[name_idx]
                        name_idx += 1
        
        # Translate series_names list
        if "series_names" in new_chart and new_chart["series_names"]:
            new_chart["series_names"] = self.translate_batch(new_chart["series_names"])
        
        # Translate categories if they are text (not numbers)
        if "categories" in new_chart and new_chart["categories"]:
            # Check if categories are text (strings)
            text_categories = [cat for cat in new_chart["categories"] if isinstance(cat, str)]
            if text_categories:
                translated_cats = self.translate_batch(new_chart["categories"])
                new_chart["categories"] = translated_cats
        
        return new_chart
    
    def translate_smartart(self, smartart: Dict) -> Dict:
        """
        Translate SmartArt text while preserving hierarchical structure.
        
        Args:
            smartart: SmartArt dictionary
            
        Returns:
            SmartArt dictionary with translated text
        """
        new_smartart = deepcopy(smartart)
        
        # Translate texts list
        if "texts" in new_smartart and new_smartart["texts"]:
            new_smartart["texts"] = self.translate_batch(new_smartart["texts"])
        
        # Translate node texts
        if "nodes" in new_smartart and new_smartart["nodes"]:
            node_texts = [node.get("text", "") for node in new_smartart["nodes"]]
            if node_texts:
                translated_node_texts = self.translate_batch(node_texts)
                for idx, node in enumerate(new_smartart["nodes"]):
                    if node.get("text"):
                        node["text"] = translated_node_texts[idx]
        
        # Update full_text
        if "texts" in new_smartart:
            new_smartart["full_text"] = " ".join(new_smartart["texts"])
        
        return new_smartart
    
    def translate_speaker_notes(self, notes: Dict) -> Dict:
        """
        Translate speaker notes while preserving metadata.
        
        Args:
            notes: Speaker notes dictionary
            
        Returns:
            Speaker notes dictionary with translated text
        """
        new_notes = deepcopy(notes)
        
        if "text" in new_notes and new_notes["text"]:
            translated = self.translate_batch([new_notes["text"]])
            new_notes["text"] = translated[0]
        
        return new_notes
    
    def translate_slide(self, slide: Dict, slide_num: int) -> Dict:
        """
        Translate a single slide while preserving all metadata including:
        - layout_info (NEW)
        - background (NEW)
        - fill, line, shadow for each element (NEW)
        - placeholder_info (NEW)
        
        Args:
            slide: Slide dictionary
            slide_num: Slide number for progress reporting
            
        Returns:
            Slide dictionary with translated content
        """
        print(f"Translating slide {slide_num}...", end=" ", flush=True)
        
        # Small delay to avoid rate limits
        if slide_num > 1:
            time.sleep(0.2)
        
        new_slide = deepcopy(slide)
        
        # Preserve layout_info, background - these don't need translation
        # They are already in new_slide via deepcopy
        
        # Translate elements
        if "elements" in new_slide:
            translated_elements = []
            for element in new_slide["elements"]:
                element_type = element.get("element_type")
                
                # Handle different element types
                if element_type == "Table":
                    # Table has table_data field
                    if "table_data" in element:
                        element["table_data"] = self.translate_table(element["table_data"])
                    translated_elements.append(element)
                    
                elif element_type == "Chart":
                    # Chart has chart_data field
                    if "chart_data" in element:
                        element["chart_data"] = self.translate_chart(element["chart_data"])
                    translated_elements.append(element)
                    
                elif element_type in ["TextBox", "AutoShape"]:
                    # These have paragraphs
                    translated_elements.append(self.translate_text_element(element))
                    
                else:
                    # Picture, Other types - preserve as is
                    translated_elements.append(deepcopy(element))
            
            new_slide["elements"] = translated_elements
        
        # Translate speaker notes
        if "speaker_notes" in new_slide and new_slide["speaker_notes"]:
            new_slide["speaker_notes"] = self.translate_speaker_notes(new_slide["speaker_notes"])
        
        # Translate SmartArt
        if "smartart" in new_slide and new_slide["smartart"]:
            translated_smartart = []
            for smartart in new_slide["smartart"]:
                translated_smartart.append(self.translate_smartart(smartart))
            new_slide["smartart"] = translated_smartart
        
        # Preserve links as is (URLs don't need translation)
        # Preserve background, layout_info (already done via deepcopy)
        
        print("✓")
        return new_slide
    
    def translate_presentation(self, input_path: str, output_path: str) -> Dict:
        """
        Translate entire presentation while preserving all metadata including:
        - slide_masters (NEW - preserved, not translated)
        - All layout information
        - Fill, line, shadow properties
        - Background information
        
        Args:
            input_path: Path to input JSON file
            output_path: Path to output JSON file
            
        Returns:
            Dictionary with translation statistics
        """
        print(f"Loading presentation from {input_path}...")
        with open(input_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        print(f"Translating to {self.target_language}...")
        print(f"Total slides: {data['total_slides']}")
        if 'slide_masters' in data:
            print(f"Slide masters: {len(data['slide_masters'])}")
        print("=" * 80)
        
        # Create new data structure preserving top-level metadata
        translated_data = {
            "presentation_name": data["presentation_name"],
            "total_slides": data["total_slides"],
            "slides": []
        }
        
        # Preserve slide_masters (NEW - these don't need translation, just structure info)
        if "slide_masters" in data:
            translated_data["slide_masters"] = deepcopy(data["slide_masters"])
        
        # Translate each slide
        start_time = time.time()
        for idx, slide in enumerate(data["slides"], 1):
            translated_slide = self.translate_slide(slide, idx)
            translated_data["slides"].append(translated_slide)
        
        elapsed_time = time.time() - start_time
        
        # Save translated data
        print("=" * 80)
        print(f"Saving translated presentation to {output_path}...")
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(translated_data, f, indent=2, ensure_ascii=False)
        
        # Print statistics
        print("\n" + "=" * 80)
        print("TRANSLATION COMPLETE!")
        print("=" * 80)
        print(f"Target language: {self.target_language}")
        print(f"Total slides translated: {data['total_slides']}")
        print(f"Total texts translated: {self.stats['total_texts_translated']}")
        print(f"API calls made: {self.stats['api_calls']}")
        print(f"Total tokens used: {self.stats['total_tokens_used']}")
        print(f"Time elapsed: {elapsed_time:.2f} seconds")
        print(f"Output saved to: {output_path}")
        print("=" * 80)
        
        return self.stats


def main():
    """Main function to run translation"""
    import argparse
    
    parser = argparse.ArgumentParser(description="Translate PowerPoint extracted content")
    parser.add_argument("input_file", help="Input JSON file path")
    parser.add_argument("-o", "--output", help="Output JSON file path (default: input_file with _translated suffix)")
    parser.add_argument("-l", "--language", default="Spanish", help="Target language (default: Spanish)")
    parser.add_argument("-k", "--api-key", help="OpenAI API key (default: from .env)")
    
    args = parser.parse_args()
    
    # Determine output path
    if args.output:
        output_path = args.output
    else:
        base_name = args.input_file.replace(".json", "")
        output_path = f"{base_name}_translated_{args.language.lower()}.json"
    
    # Create translator
    translator = PPTTranslator(api_key=args.api_key, target_language=args.language)
    
    # Translate
    stats = translator.translate_presentation(args.input_file, output_path)
    
    return stats


if __name__ == "__main__":
    main()