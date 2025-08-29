import streamlit as st
import os
import ppt_generator
from dotenv import load_dotenv

load_dotenv()

# Setting up the page layout done here
st.set_page_config(
    page_title="PowerPoint Presentation Generator",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("💡 AI-Powered Presentation Generator")
st.markdown("Enter a topic below to auto-generate a professional PowerPoint presentation.")
st.markdown("---")

# Users input 
topic = st.text_input("Enter your presentation topic:", placeholder="e.g., The future of renewable energy")

if st.button("Generate Presentation", help="Click to start the generation process."):
    if not topic:
        st.error("Please enter a topic to continue.")
    else:
        with st.spinner(f"Generating presentation on '{topic}'... This may take a few moments."):
            try:
                # instance created
                generator = ppt_generator.PowerPointGenerator()
                
                # Called the refactored method with the user's topic
                filename = generator.generate_presentation(topic=topic)
                
                if filename:
                    st.success("Presentation generated successfully!")
                    st.balloons()

                    # Provided a download link for the file
                    with open(filename, "rb") as file:
                        st.download_button(
                            label="📥 Download PowerPoint File",
                            data=file,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                    st.info(f"The file is also saved in your project folder as: **{filename}**")
                else:
                    st.error("Presentation generation failed. Please check the logs.")

            except ValueError as ve:
                st.error(f"Configuration Error: {ve}")
            except Exception as e:
                st.error(f"An unexpected error occurred: {e}")

st.markdown("---")
st.info("To run this app, make sure your virtual environment is active, then run `streamlit run app.py` in your terminal.")
