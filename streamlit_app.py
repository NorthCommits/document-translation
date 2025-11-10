"""
Streamlit App for PowerPoint Translation Pipeline
=================================================
A user-friendly web interface for translating PowerPoint presentations.
"""

import streamlit as st
import os
import time
import tempfile
from datetime import datetime
import sys

# Import the pipeline classes
from ppt_translation_assembly_pipeline import PPTXExtractor, PPTTranslator, PPTXReassembler


# Page configuration
st.set_page_config(
    page_title="PPT Translation",
    layout="centered"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main-header {
        font-size: 3rem;
        font-weight: bold;
        text-align: center;
        color: #1f77b4;
        margin-bottom: 2rem;
    }
    .progress-box {
        padding: 1.5rem;
        border-radius: 10px;
        background-color: #f0f2f6;
        margin: 1rem 0;
    }
    .stage-header {
        font-size: 1.2rem;
        font-weight: bold;
        margin-bottom: 0.5rem;
    }
    .time-info {
        color: #666;
        font-size: 0.9rem;
    }
    .success-box {
        padding: 1rem;
        border-radius: 10px;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

# Main header
st.markdown('<h1 class="main-header">PPT Translation</h1>', unsafe_allow_html=True)

# Initialize session state
if 'extraction_done' not in st.session_state:
    st.session_state.extraction_done = False
    st.session_state.translation_done = False
    st.session_state.reassembly_done = False
    st.session_state.extraction_time = 0
    st.session_state.translation_time = 0
    st.session_state.reassembly_time = 0
    st.session_state.output_file = None
    st.session_state.processing = False

# Create two columns for upload and language selection
col1, col2 = st.columns([2, 1])

with col1:
    # File upload
    uploaded_file = st.file_uploader(
        "Upload PowerPoint Presentation",
        type=['ppt', 'pptx'],
        help="Upload a .ppt or .pptx file to translate (Max 20MB)"
    )

with col2:
    # Language selection dropdown
    target_language = st.selectbox(
        "Target Language",
        ["Spanish", "French", "German", "Italian", "Portuguese", "Dutch", 
         "Polish", "Swedish", "Danish", "Norwegian", "Finnish", "Greek",
         "Czech", "Hungarian", "Romanian", "Bulgarian", "Croatian", 
         "Japanese", "Chinese"],
        index=0
    )

# Display file info if uploaded and check size
if uploaded_file is not None:
    file_size = uploaded_file.size / (1024 * 1024)  # Convert to MB
    
    # Check if file exceeds 20MB
    if file_size > 20:
        st.error(f"❌ File size ({file_size:.2f} MB) exceeds the 20MB limit. Please upload a smaller file.")
        uploaded_file = None
    else:
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"📄 **File:** {uploaded_file.name}")
        with col2:
            st.info(f"💾 **Size:** {file_size:.2f} MB")

# Start translation button
if uploaded_file is not None and not st.session_state.processing:
    if st.button("🚀 Start Translation", type="primary", use_container_width=True):
        st.session_state.processing = True
        st.session_state.extraction_done = False
        st.session_state.translation_done = False
        st.session_state.reassembly_done = False
        st.rerun()

# Processing pipeline
if st.session_state.processing and uploaded_file is not None:
    try:
        # Create temporary directory for processing
        with tempfile.TemporaryDirectory() as temp_dir:
            # Save uploaded file
            input_path = os.path.join(temp_dir, uploaded_file.name)
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            base_name = os.path.splitext(uploaded_file.name)[0]
            extracted_json = os.path.join(temp_dir, f"{base_name}_extracted.json")
            translated_json = os.path.join(temp_dir, f"{base_name}_translated.json")
            output_pptx = os.path.join(temp_dir, f"{base_name}_{target_language.lower()}.pptx")
            
            # Progress container
            progress_container = st.container()
            
            with progress_container:
                st.markdown("### 🔄 Translation Progress")
                
                # Stage 1: Extraction
                if not st.session_state.extraction_done:
                    st.markdown("#### 📤 Stage 1: Extraction")
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    status_text.text("Extracting content from PowerPoint...")
                    start_time = time.time()
                    
                    try:
                        extractor = PPTXExtractor(input_path)
                        progress_bar.progress(30)
                        
                        extracted_data = extractor.extract_all()
                        progress_bar.progress(70)
                        
                        extractor.save_to_json(extracted_json)
                        progress_bar.progress(100)
                        
                        st.session_state.extraction_time = time.time() - start_time
                        st.session_state.extraction_done = True
                        
                        st.success(f"✅ Extraction completed in {st.session_state.extraction_time:.2f} seconds")
                        st.markdown(f"**Slides extracted:** {len(extracted_data['slides'])}")
                        
                    except Exception as e:
                        st.error(f"❌ Extraction failed: {str(e)}")
                        st.session_state.processing = False
                        st.stop()
                else:
                    st.success(f"✅ Extraction completed in {st.session_state.extraction_time:.2f} seconds")
                
                # Stage 2: Translation
                if st.session_state.extraction_done and not st.session_state.translation_done:
                    st.markdown("#### 🌐 Stage 2: Translation")
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    status_text.text(f"Translating to {target_language}...")
                    start_time = time.time()
                    
                    try:
                        # Use API key from .env file only
                        translator = PPTTranslator(api_key=None, target_language=target_language)
                        progress_bar.progress(20)
                        
                        # Create a custom progress callback
                        import json
                        with open(extracted_json, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                        
                        total_slides = len(data['slides'])
                        
                        status_text.text(f"Translating {total_slides} slides to {target_language}...")
                        progress_bar.progress(40)
                        
                        translation_stats = translator.translate_presentation(extracted_json, translated_json)
                        progress_bar.progress(100)
                        
                        st.session_state.translation_time = time.time() - start_time
                        st.session_state.translation_done = True
                        
                        st.success(f"✅ Translation completed in {st.session_state.translation_time:.2f} seconds")
                        st.markdown(f"**Texts translated:** {translation_stats['total_texts_translated']}")
                        st.markdown(f"**API calls:** {translation_stats['api_calls']}")
                        st.markdown(f"**Tokens used:** {translation_stats['total_tokens_used']}")
                        
                    except Exception as e:
                        st.error(f"❌ Translation failed: {str(e)}")
                        st.session_state.processing = False
                        st.stop()
                else:
                    if st.session_state.translation_done:
                        st.success(f"✅ Translation completed in {st.session_state.translation_time:.2f} seconds")
                
                # Stage 3: Reassembly
                if st.session_state.translation_done and not st.session_state.reassembly_done:
                    st.markdown("#### 🔧 Stage 3: Reassembly")
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    status_text.text("Reassembling PowerPoint presentation...")
                    start_time = time.time()
                    
                    try:
                        reassembler = PPTXReassembler(input_path, translated_json)
                        progress_bar.progress(30)
                        
                        reassembly_stats = reassembler.reassemble(output_pptx)
                        progress_bar.progress(100)
                        
                        st.session_state.reassembly_time = time.time() - start_time
                        st.session_state.reassembly_done = True
                        
                        # Read the output file and store in session state
                        with open(output_pptx, "rb") as f:
                            st.session_state.output_file = f.read()
                        
                        st.session_state.output_filename = f"{base_name}_{target_language.lower()}.pptx"
                        
                        st.success(f"✅ Reassembly completed in {st.session_state.reassembly_time:.2f} seconds")
                        st.markdown(f"**Slides processed:** {reassembly_stats['slides_processed']}")
                        st.markdown(f"**Elements updated:** {reassembly_stats['elements_updated']}")
                        
                    except Exception as e:
                        st.error(f"❌ Reassembly failed: {str(e)}")
                        st.session_state.processing = False
                        st.stop()
                else:
                    if st.session_state.reassembly_done:
                        st.success(f"✅ Reassembly completed in {st.session_state.reassembly_time:.2f} seconds")
            
            # Mark processing as complete
            st.session_state.processing = False
            
    except Exception as e:
        st.error(f"❌ An error occurred: {str(e)}")
        st.session_state.processing = False

# Summary and download section
if st.session_state.reassembly_done and st.session_state.output_file is not None:
    st.markdown("---")
    st.markdown("### 🎉 Translation Complete!")
    
    # Summary box
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Extraction Time", f"{st.session_state.extraction_time:.2f}s")
    with col2:
        st.metric("Translation Time", f"{st.session_state.translation_time:.2f}s")
    with col3:
        st.metric("Reassembly Time", f"{st.session_state.reassembly_time:.2f}s")
    
    total_time = st.session_state.extraction_time + st.session_state.translation_time + st.session_state.reassembly_time
    st.info(f"⏱️ **Total Processing Time:** {total_time:.2f} seconds")
    
    # Download button
    st.download_button(
        label="📥 Download Translated PowerPoint",
        data=st.session_state.output_file,
        file_name=st.session_state.output_filename,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
        use_container_width=True
    )
    
    # Reset button
    if st.button("🔄 Translate Another File", use_container_width=True):
        st.session_state.extraction_done = False
        st.session_state.translation_done = False
        st.session_state.reassembly_done = False
        st.session_state.extraction_time = 0
        st.session_state.translation_time = 0
        st.session_state.reassembly_time = 0
        st.session_state.output_file = None
        st.session_state.processing = False
        st.rerun()

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 2rem 0;'>
    <p>Powered by OpenAI GPT-4o-mini | Built with Streamlit</p>
</div>
""", unsafe_allow_html=True)