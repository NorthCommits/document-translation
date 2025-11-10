import streamlit as st
import os
import time
import tempfile
import json

from extractor import PPTXExtractor
from translator import PPTTranslator
from reassembler import PPTXReassembler

st.set_page_config(
    page_title="PPT Translation",
    layout="centered"
)

st.markdown("""
    <style>
    .main-header {
        font-size: 3rem;
        font-weight: bold;
        text-align: center;
        color: #1f77b4;
        margin-bottom: 2rem;
    }
    .stage-complete {
        color: #28a745;
        font-weight: bold;
    }
    .stage-processing {
        color: #007bff;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="main-header">PPT Translation</h1>', unsafe_allow_html=True)

if 'extraction_done' not in st.session_state:
    st.session_state.extraction_done = False
    st.session_state.translation_done = False
    st.session_state.reassembly_done = False
    st.session_state.extraction_time = 0
    st.session_state.translation_time = 0
    st.session_state.reassembly_time = 0
    st.session_state.output_file = None
    st.session_state.processing = False

col1, col2 = st.columns([2, 1])

with col1:
    uploaded_file = st.file_uploader(
        "Upload PowerPoint Presentation",
        type=['ppt', 'pptx'],
        help="Upload a .ppt or .pptx file to translate (Max 20MB)"
    )

with col2:
    target_language = st.selectbox(
        "Target Language",
        ["Spanish", "French", "German", "Italian", "Portuguese", "Dutch", 
         "Polish", "Swedish", "Danish", "Norwegian", "Finnish", "Greek",
         "Czech", "Hungarian", "Romanian", "Bulgarian", "Croatian", 
         "Japanese", "Chinese"],
        index=0
    )

if uploaded_file is not None:
    file_size = uploaded_file.size / (1024 * 1024)
    
    if file_size > 20:
        st.error(f"File size ({file_size:.2f} MB) exceeds the 20MB limit. Please upload a smaller file.")
        uploaded_file = None
    else:
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"File: {uploaded_file.name}")
        with col2:
            st.info(f"Size: {file_size:.2f} MB")

if uploaded_file is not None and not st.session_state.processing:
    if st.button("Start Translation", type="primary", use_container_width=True):
        st.session_state.processing = True
        st.session_state.extraction_done = False
        st.session_state.translation_done = False
        st.session_state.reassembly_done = False
        st.rerun()

if st.session_state.processing and uploaded_file is not None:
    try:
        # Create json_bin directory in project root if it doesn't exist
        json_bin_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "json_bin")
        os.makedirs(json_bin_dir, exist_ok=True)
        
        # Save uploaded file to json_bin directory
        base_name = os.path.splitext(uploaded_file.name)[0]
        input_path = os.path.join(json_bin_dir, uploaded_file.name)
        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Define paths for intermediate files in json_bin
        extracted_json = os.path.join(json_bin_dir, f"{base_name}_extracted.json")
        translated_json = os.path.join(json_bin_dir, f"{base_name}_translated.json")
        output_pptx = os.path.join(json_bin_dir, f"{base_name}_{target_language.lower()}.pptx")
        
        progress_container = st.container()
        
        with progress_container:
            st.markdown("### Translation Progress")
            
            if not st.session_state.extraction_done:
                st.markdown("#### Stage 1: Extraction")
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                status_text.text("Extracting content from PowerPoint...")
                start_time = time.time()
                
                try:
                    extractor = PPTXExtractor(input_path)
                    progress_bar.progress(30)
                    status_text.text("Extracting slides and content...")
                    
                    extracted_data = extractor.extract_all()
                    progress_bar.progress(70)
                    status_text.text("Saving extraction data...")
                    
                    extractor.save_to_json(extracted_json)
                    progress_bar.progress(100)
                    
                    st.session_state.extraction_time = time.time() - start_time
                    st.session_state.extraction_done = True
                    
                    st.success(f"Extraction completed in {st.session_state.extraction_time:.2f} seconds")
                    st.markdown(f"**Slides extracted:** {len(extracted_data['slides'])}")
                    
                except Exception as e:
                    st.error(f"Extraction failed: {str(e)}")
                    st.session_state.processing = False
                    st.stop()
            else:
                st.success(f"Extraction completed in {st.session_state.extraction_time:.2f} seconds")
            
            if st.session_state.extraction_done and not st.session_state.translation_done:
                st.markdown("#### Stage 2: Translation")
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                status_text.text(f"Translating to {target_language}...")
                start_time = time.time()
                
                try:
                    translator = PPTTranslator(api_key=None, target_language=target_language)
                    progress_bar.progress(10)
                    
                    with open(extracted_json, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    
                    total_slides = len(data['slides'])
                    status_text.text(f"Translating {total_slides} slides...")
                    progress_bar.progress(20)
                    
                    translation_stats = translator.translate_presentation(extracted_json, translated_json)
                    progress_bar.progress(100)
                    
                    st.session_state.translation_time = time.time() - start_time
                    st.session_state.translation_done = True
                    
                    st.success(f"Translation completed in {st.session_state.translation_time:.2f} seconds")
                    st.markdown(f"**Texts translated:** {translation_stats['total_texts_translated']}")
                    st.markdown(f"**API calls:** {translation_stats['api_calls']}")
                    st.markdown(f"**Tokens used:** {translation_stats['total_tokens_used']:,}")
                    st.markdown(f"**Total cost:** ${translation_stats['total_cost_usd']:.4f} USD")
                    
                except Exception as e:
                    st.error(f"Translation failed: {str(e)}")
                    st.session_state.processing = False
                    st.stop()
            else:
                if st.session_state.translation_done:
                    st.success(f"Translation completed in {st.session_state.translation_time:.2f} seconds")
            
            if st.session_state.translation_done and not st.session_state.reassembly_done:
                st.markdown("#### Stage 3: Reassembly")
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                status_text.text("Reassembling PowerPoint presentation...")
                start_time = time.time()
                
                try:
                    reassembler = PPTXReassembler(input_path, translated_json)
                    progress_bar.progress(30)
                    status_text.text("Applying translations...")
                    
                    reassembly_stats = reassembler.reassemble(output_pptx)
                    progress_bar.progress(100)
                    
                    st.session_state.reassembly_time = time.time() - start_time
                    st.session_state.reassembly_done = True
                    
                    with open(output_pptx, "rb") as f:
                        st.session_state.output_file = f.read()
                    
                    st.session_state.output_filename = f"{base_name}_{target_language.lower()}.pptx"
                    
                    st.success(f"Reassembly completed in {st.session_state.reassembly_time:.2f} seconds")
                    st.markdown(f"**Slides processed:** {reassembly_stats['slides_processed']}")
                    st.markdown(f"**Elements updated:** {reassembly_stats['elements_updated']}")
                    
                except Exception as e:
                    st.error(f"Reassembly failed: {str(e)}")
                    st.session_state.processing = False
                    st.stop()
            else:
                if st.session_state.reassembly_done:
                    st.success(f"Reassembly completed in {st.session_state.reassembly_time:.2f} seconds")
        
        st.session_state.processing = False
        
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
        st.session_state.processing = False

if st.session_state.reassembly_done and st.session_state.output_file is not None:
    st.markdown("---")
    st.markdown("### Translation Complete")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Extraction Time", f"{st.session_state.extraction_time:.2f}s")
    with col2:
        st.metric("Translation Time", f"{st.session_state.translation_time:.2f}s")
    with col3:
        st.metric("Reassembly Time", f"{st.session_state.reassembly_time:.2f}s")
    
    total_time = st.session_state.extraction_time + st.session_state.translation_time + st.session_state.reassembly_time
    st.info(f"Total Processing Time: {total_time:.2f} seconds")
    
    st.download_button(
        label="Download Translated PowerPoint",
        data=st.session_state.output_file,
        file_name=st.session_state.output_filename,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
        use_container_width=True
    )
    
    if st.button("Translate Another File", use_container_width=True):
        st.session_state.extraction_done = False
        st.session_state.translation_done = False
        st.session_state.reassembly_done = False
        st.session_state.extraction_time = 0
        st.session_state.translation_time = 0
        st.session_state.reassembly_time = 0
        st.session_state.output_file = None
        st.session_state.processing = False
        st.rerun()

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 2rem 0;'>
    <p>Powered by OpenAI GPT-4o-mini | Built with Streamlit</p>
</div>
""", unsafe_allow_html=True)