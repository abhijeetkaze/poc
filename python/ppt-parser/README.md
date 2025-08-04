# ppt-parser PoC

This project is a proof of concept (PoC) for parsing PowerPoint (PPT) files in Python, with the goal of capturing slides as images.

## Current Status

- **Image Extraction:**  
    Capturing slides as images directly from PPT files is **not supported in Python** due to the closed-source nature of the PPT file format by Microsoft.
- **Text Extraction:**  
    Extracting text content from slides is possible using available Python libraries.
- **Alternative Approach:**  
    As a workaround, you can use the LibreOffice command-line interface (CLI) to convert or export slides as images.

## Summary

- This PoC demonstrates the limitations of Python libraries for PPT image extraction.
- For image export, consider automating LibreOffice or similar tools outside of Python.