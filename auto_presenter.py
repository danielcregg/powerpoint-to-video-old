from dotenv import load_dotenv
import os
import sys
import subprocess # New import for running command-line tools
import time
import google.generativeai as genai
from TTS.api import TTS # Using the high-quality offline TTS
from moviepy.editor import ImageClip, AudioFileClip, concatenate_videoclips
import fitz # PyMuPDF

# --- CONFIGURATION ---
load_dotenv()
# Your Gemini API Key for script generation
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")


def configure_gemini_vision_model(api_key):
    """Configures and returns the Gemini vision model."""
    print("--- Configuring Gemini Vision Model ---")
    if not api_key or api_key == "YOUR_GEMINI_API_KEY":
        print("Error: Please paste your Gemini API key into the GEMINI_API_KEY variable.")
        sys.exit(1)
    try:
        genai.configure(api_key=api_key)
        
        # Get all available models that support vision/content generation
        available_models = []
        print("  - Available Gemini models:")
        for model in genai.list_models():
            if 'generateContent' in model.supported_generation_methods:
                available_models.append(model.name)
                # Show model info if available
                display_name = model.display_name if hasattr(model, 'display_name') else model.name
                print(f"    • {model.name} ({display_name})")
        
        print(f"\n  - Found {len(available_models)} models with vision/content generation support")
        
        # Priority order for script generation - optimized for high request volume and quality
        model_priorities = [
            # Gemini 2.5 models (newest, most capable)
            "models/gemini-2.5-flash",           # Latest Flash - best balance of speed/quality
            "models/gemini-2.5-flash-lite-preview-06-17",  # Most cost-efficient, high throughput
            "models/gemini-2.5-pro",             # Most capable but may have lower limits
            
            # Gemini 2.0 models
            "models/gemini-2.0-flash",           # Next generation features
            "models/gemini-2.0-flash-lite",     # Cost efficient with low latency
            
            # Gemini 1.5 models (proven and reliable)
            "models/gemini-1.5-flash",           # Fast and versatile
            "models/gemini-1.5-flash-8b",       # High volume, lower intelligence tasks
            "models/gemini-1.5-flash-latest",   # Latest 1.5 Flash
            "models/gemini-1.5-pro",            # Complex reasoning (but lower rate limits)
        ]
        
        # Select the best available model based on priority
        selected_model = None
        for preferred_model in model_priorities:
            if preferred_model in available_models:
                selected_model = preferred_model
                print(f"  - ✓ Selected: {selected_model}")
                break
        
        if not selected_model:
            # Fallback: use any model with key terms, prioritizing newer versions
            fallback_keywords = ['2.5-flash', '2.0-flash', '1.5-flash', 'flash', 'pro']
            for keyword in fallback_keywords:
                for model_name in available_models:
                    if keyword in model_name.lower():
                        selected_model = model_name
                        print(f"  - ✓ Fallback selected: {selected_model}")
                        break
                if selected_model:
                    break
        
        if not selected_model:
            print("Error: Could not find a suitable Gemini model for script generation.")
            print(f"Available models: {available_models}")
            sys.exit(1)
        
        # Provide information about the selected model
        model_info = ""
        if "2.5" in selected_model:
            if "flash-lite" in selected_model:
                model_info = "(Gemini 2.5 Flash Lite - Most cost-efficient, high throughput)"
            elif "flash" in selected_model:
                model_info = "(Gemini 2.5 Flash - Latest with adaptive thinking)"
            elif "pro" in selected_model:
                model_info = "(Gemini 2.5 Pro - Enhanced reasoning, may have lower rate limits)"
        elif "2.0" in selected_model:
            if "flash-lite" in selected_model:
                model_info = "(Gemini 2.0 Flash Lite - Cost efficient, low latency)"
            elif "flash" in selected_model:
                model_info = "(Gemini 2.0 Flash - Next generation features)"
        elif "1.5" in selected_model:
            if "flash-8b" in selected_model:
                model_info = "(Gemini 1.5 Flash-8B - Optimized for high volume tasks)"
            elif "flash" in selected_model:
                model_info = "(Gemini 1.5 Flash - Fast and versatile)"
            elif "pro" in selected_model:
                model_info = "(Gemini 1.5 Pro - Complex reasoning, lower rate limits)"
        
        print(f"  - Model Type: {model_info}")
        print(f"  - Perfect for batch processing PowerPoint presentations!")
        
        return genai.GenerativeModel(model_name=selected_model)
    except Exception as e:
        print(f"An error occurred during Gemini configuration: {e}")
        sys.exit(1)

# --- NEW LINUX-COMPATIBLE FUNCTION ---
def extract_slides_as_images_linux(pptx_path, temp_folder):
    """
    Converts PPTX slides to PNG images using LibreOffice on Linux.
    This replaces the PowerPoint dependency.
    """
    print("\nStep 1: Converting PPTX to images (using LibreOffice for PDF export)...")
    if not os.path.exists(temp_folder):
        os.makedirs(temp_folder)
    
    try:
        # Construct the command to run LibreOffice in headless mode
        command = [
            "soffice", # The command for LibreOffice
            "--headless",
            "--convert-to", "pdf",
            "--outdir", temp_folder,
            pptx_path
        ]
        print(f"  - Running command: {' '.join(command)}")
        # Execute the command
        subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        print(f"  - Successfully converted PPTX to PDF using LibreOffice.")
    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        print(f"  - Error during PDF conversion with LibreOffice: {e}")
        print("  - Ensure LibreOffice is installed in your Codespace environment.")
        return None

    # This part remains the same, as PyMuPDF is cross-platform
    pdf_filename = os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf"
    pdf_path = os.path.join(temp_folder, pdf_filename)
    
    image_paths = []
    
    # Temporarily suppress MuPDF warnings about interactive elements
    print("  - Extracting slide images from PDF...")
    original_stderr = sys.stderr
    try:
        # Redirect stderr to suppress MuPDF warnings about Screen annotations
        import io
        sys.stderr = io.StringIO()
        
        doc = fitz.open(pdf_path)
        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=300)
            image_path = os.path.join(temp_folder, f"slide_{i + 1}.png")
            pix.save(image_path)
            image_paths.append(image_path)
        doc.close()
        
    finally:
        # Restore stderr
        sys.stderr = original_stderr
    
    print(f"  - Successfully extracted {len(image_paths)} slide images")
    return image_paths

def generate_script_for_slide(vision_model, image_path, slide_number, total_slides):
    """Generates a speaker script for a slide image using Gemini."""
    print(f"\nStep 2: Generating script for slide {slide_number} (using Gemini)...")
    try:
        slide_image = genai.upload_file(image_path)
        
        # Create context-aware prompts based on slide position
        if slide_number == 1:
            context_prompt = "This is the first slide of the presentation. You may greet the audience and introduce the topic."
        elif slide_number == total_slides:
            context_prompt = "This is the final slide of the presentation. Thank the audience, summarize key takeaways, or provide a professional closing."
        else:
            context_prompt = "This is a middle slide of the presentation. Continue the presentation flow without greetings or farewells."
        
        prompt = [
            "You are a professional presenter. Write a clear and engaging speaker script for this slide.",
            context_prompt,
            "Explain the key points as if presenting to an audience.",
            "Do not describe the slide's layout. Deliver the information directly.",
            "Keep the script under 150 words.",
            slide_image
        ]
        response = vision_model.generate_content(prompt)
        script = response.text.strip().replace("*", "")
        print(f"  - Script for slide {slide_number} generated successfully.")
        genai.delete_file(slide_image.name)
        return script
    except Exception as e:
        print(f"  - Error generating script for slide {slide_number}: {e}")
        return None

def synthesize_speech_with_coqui(tts_engine, text, output_path, slide_number):
    """Converts text to a WAV audio file using the offline Coqui TTS engine."""
    print(f"Step 3: Synthesizing audio for slide {slide_number} (using local Coqui TTS)...")
    if not text:
        print("  - Skipping audio synthesis due to empty script.")
        return None
    try:
        print(f"  - Starting TTS synthesis for slide {slide_number}...")
        # Use tts method instead of to_file
        tts_engine.tts_to_file(text=text, file_path=output_path)
        print(f"  - TTS synthesis completed for slide {slide_number}")
        if os.path.exists(output_path):
            print(f"  - Audio file saved: {output_path}")
            return output_path
        else:
            print(f"  - Error: Audio file was not created at {output_path}")
            return None
    except KeyboardInterrupt:
        print(f"  - Process interrupted by user for slide {slide_number}")
        return None
    except Exception as e:
        print(f"  - Error synthesizing speech for slide {slide_number}: {e}")
        print(f"  - Error type: {type(e).__name__}")
        return None

def create_video_with_moviepy(image_files, audio_files, output_path):
    """Creates a video by combining slide images and audio narrations using moviepy."""
    print("\nStep 4: Creating video from images and audio with moviepy...")
    clips = []
    for img_path, audio_path in zip(image_files, audio_files):
        if not os.path.exists(img_path):
            print(f"  - Warning: Missing image {img_path}. Skipping slide.")
            continue
        if not audio_path or not os.path.exists(audio_path):
            print(f"  - Warning: Missing audio for {os.path.basename(img_path)}. Skipping slide.")
            continue
        try:
            audio_clip = AudioFileClip(audio_path)
            image_clip = ImageClip(img_path)
            image_clip = image_clip.set_duration(audio_clip.duration)
            video_clip = image_clip.set_audio(audio_clip)
            clips.append(video_clip)
            print(f"  - Processed slide: {os.path.basename(img_path)}")
        except Exception as e:
            print(f"  - Error processing clip for {os.path.basename(img_path)}: {e}")

    if not clips:
        print("  - No clips were created. Cannot generate video.")
        print("  - This is likely due to TTS synthesis failures. Check the TTS errors above.")
        return

    print(f"  - Created {len(clips)} video clips successfully")
    final_video = concatenate_videoclips(clips)
    
    try:
        print(f"  - Writing video file: {output_path}")
        # Use more compatible settings for containerized environments
        final_video.write_videofile(
            output_path, 
            fps=24, 
            codec='libx264',
            audio_codec='aac',
            temp_audiofile='temp-audio.m4a',
            remove_temp=True,
            verbose=False,
            logger=None
        )
        print(f"\nVideo successfully created: {output_path}")
        
        # Verify the file was created and has content
        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            print(f"  - Video file size: {os.path.getsize(output_path)} bytes")
        else:
            print("  - Warning: Video file appears to be empty or missing")
            
    except Exception as e:
        print(f"\nError writing final video file: {e}")
        
        # Remove the failed file if it exists
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
                print(f"  - Removed failed video file: {output_path}")
            except OSError:
                pass
        
        print("  - Trying alternative H.264 method...")
        
        # Try H.264 with simpler settings - overwrite original file
        try:
            final_video.write_videofile(
                output_path,
                fps=24,
                codec='libx264',
                preset='ultrafast',  # Faster encoding, larger file
                verbose=False,
                logger=None
            )
            print(f"Video successfully created with alternative H.264 settings: {output_path}")
        except Exception as e2:
            print(f"H.264 alternative failed: {e2}")
            print("  - Falling back to MP4V...")
            
            # Remove the failed file if it exists
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                    print(f"  - Removed failed video file: {output_path}")
                except OSError:
                    pass

            # Final fallback to MP4V - still use original filename
            try:
                final_video.write_videofile(
                    output_path,
                    fps=24,
                    codec='mpeg4',
                    verbose=False,
                    logger=None
                )
                print(f"Video successfully created with MP4V codec: {output_path}")
            except Exception as e3:
                print(f"All encoding methods failed: {e3}")
                print(f"Could not create video file: {output_path}")
    
    finally:
        # Clean up clips to free memory
        for clip in clips:
            clip.close()
        final_video.close()

def save_script_to_file(script, script_path, slide_number):
    """Saves the generated script to a text file."""
    try:
        with open(script_path, 'w', encoding='utf-8') as f:
            f.write(script)
        print(f"  - Script saved to: {script_path}")
        return True
    except Exception as e:
        print(f"  - Error saving script for slide {slide_number}: {e}")
        return False

def load_script_from_file(script_path):
    """Loads a script from a text file."""
    try:
        with open(script_path, 'r', encoding='utf-8') as f:
            script = f.read().strip()
        return script
    except Exception as e:
        print(f"  - Error loading script from {script_path}: {e}")
        return None

def should_regenerate_audio(script_path, audio_path):
    """Checks if audio should be regenerated based on script modification time."""
    if not os.path.exists(audio_path):
        return True  # Audio doesn't exist, need to generate
    
    if not os.path.exists(script_path):
        return False  # No script file, keep existing audio
    
    script_mtime = os.path.getmtime(script_path)
    audio_mtime = os.path.getmtime(audio_path)
    
    return script_mtime > audio_mtime  # Regenerate if script is newer

def main():
    if len(sys.argv) < 2:
        print("Usage: python auto_presenter.py <path_to_presentation.pptx>")
        print("Example: python auto_presenter.py my_presentation.pptx")
        sys.exit(1)
    
    vision_model = configure_gemini_vision_model(GEMINI_API_KEY)
    
    print("\n--- Initializing Local Coqui TTS Engine ---")
    print("This may take a moment and will download model files on the first run...")
    try:
        tts_engine = TTS("tts_models/en/ljspeech/vits")
        print("--- Coqui TTS Engine Initialized Successfully ---")
    except Exception as e:
        print(f"Error initializing Coqui TTS: {e}")
        sys.exit(1)

    input_pptx = os.path.abspath(sys.argv[1])
    
    # Better file validation
    if not os.path.exists(input_pptx):
        print(f"Error: File not found at {input_pptx}")
        print("Please check the file path and ensure the file exists.")
        
        # Check if user provided a filename without extension
        if not input_pptx.endswith('.pptx'):
            suggested_path = input_pptx + '.pptx'
            if os.path.exists(suggested_path):
                print(f"Did you mean: {suggested_path}?")
            else:
                print("Note: The file should have a .pptx extension")
        sys.exit(1)
    
    # Check file extension
    if not input_pptx.lower().endswith('.pptx'):
        print(f"Error: Expected a PowerPoint file (.pptx), but got: {input_pptx}")
        print("This script only works with PowerPoint (.pptx) files.")
        sys.exit(1)
        
    base_dir = os.path.dirname(input_pptx)
    file_name = os.path.splitext(os.path.basename(input_pptx))[0]
    temp_dir = os.path.join(base_dir, f"{file_name}_temp_files")
    
    # Call the new Linux-compatible function
    slide_images = extract_slides_as_images_linux(input_pptx, temp_dir)
    if not slide_images:
        sys.exit(1)

    audio_files = []
    successful_audio_count = 0
    
    print(f"\n--- Processing {len(slide_images)} slides ---")
    print("Note: You can edit script files in the temp folder and rerun to regenerate audio for modified scripts.")
    
    for i, img_path in enumerate(slide_images):
        slide_num = i + 1
        audio_path = os.path.join(temp_dir, f"audio_{slide_num}.wav")
        script_path = os.path.join(temp_dir, f"script_{slide_num}.txt")

        # Load existing script or generate new one
        script = None
        if os.path.exists(script_path):
            print(f"\n--- Loading existing script for slide {slide_num} ---")
            script = load_script_from_file(script_path)
            if script:
                print(f"  - Script loaded from: {script_path}")
                print(f"  - Script preview: {script[:100]}..." if len(script) > 100 else f"  - Script: {script}")
            else:
                print(f"  - Failed to load script, will generate new one")
        
        if not script:
            script = generate_script_for_slide(vision_model, img_path, slide_num, len(slide_images))
            if script:
                save_script_to_file(script, script_path, slide_num)
                print(f"  - Generated script: {script[:100]}..." if len(script) > 100 else f"  - Generated script: {script}")

        # Check if we need to regenerate audio
        if script:
            if should_regenerate_audio(script_path, audio_path):
                if os.path.exists(audio_path):
                    print(f"  - Script modified, regenerating audio for slide {slide_num}")
                else:
                    print(f"Step 3: Synthesizing audio for slide {slide_num} (using local Coqui TTS)...")
                
                synthesized_audio = synthesize_speech_with_coqui(tts_engine, script, audio_path, slide_num)
                if synthesized_audio:
                    successful_audio_count += 1
                audio_files.append(synthesized_audio)
            else:
                print(f"\n--- Audio for slide {slide_num} is up to date. Skipping synthesis. ---")
                audio_files.append(audio_path)
                successful_audio_count += 1
        else:
            print(f"  - No script available for slide {slide_num}")
            audio_files.append(None)

    print(f"\n--- Audio Generation Summary ---")
    print(f"  - Total slides: {len(slide_images)}")
    print(f"  - Successful audio files: {successful_audio_count}")
    print(f"  - Failed audio files: {len(slide_images) - successful_audio_count}")
    
    if successful_audio_count > 0:
        print(f"\n--- Script Files Location ---")
        print(f"  - Script files saved in: {temp_dir}")
        print(f"  - You can edit script_*.txt files and rerun to regenerate audio")

    if successful_audio_count == 0:
        print("  - No audio files were created. Cannot generate video.")
        sys.exit(1)

    video_output_path = os.path.abspath(os.path.join(base_dir, f"{file_name}_presentation.mp4"))
    print(f"\n--- Starting Video Creation ---")
    create_video_with_moviepy(slide_images, audio_files, video_output_path)
        
    print("\nProcess finished successfully!")

if __name__ == "__main__":
    main()
