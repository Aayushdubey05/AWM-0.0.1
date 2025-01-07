import serial
import time 
import os
from docx2pdf import convert
import subprocess
import logging
from pathlib import Path    # for checking if file exists
import vpype  
from PIL import Image
import numpy as np
import cv2
from pdf2image import convert_from_path

class getDocs:
    def __init__(self):
        #self.ser = serial.Serial('/dev/ttyACM0', 9600)
        self.logger = logging.getLogger(__name__)
        logging.basicConfig(filename='getDocs.log', level=logging.DEBUG)
        #time.sleep(2)
    
    def convert_docx_to_pdf(self, docx_path):
        try:
            pdf_path  = str(Path(docx_path).with_suffix('.pdf'))
            convert(docx_path, pdf_path)
            self.logger.info(f"Converted {docx_path} to {pdf_path}")
            return pdf_path
        except Exception as e:
            self.logger.error(f"Error converting {docx_path} to pdf: {str(e)}")
            raise e
        
    def convert_to_svg(self, input_path):
        try:
            # First try using Inkscape
            self.logger.info("Attempting conversion with Inkscape...")
            return self._convert_with_inkscape(input_path)
        except Exception as e:
            self.logger.warning(f"Inkscape conversion failed: {str(e)}")
            try:
                # Fallback to OpenCV + manual SVG creation
                self.logger.info("Attempting conversion with OpenCV...")
                return self._convert_with_opencv(input_path)
            except Exception as e2:
                self.logger.error(f"All conversion methods failed: {str(e2)}")
                raise ValueError(f"Failed to convert {input_path} to SVG. Please ensure Inkscape is installed or try a different image.")
            

    def _convert_with_inkscape(self, input_path):
        svg_path = str(Path(input_path).with_suffix('.svg'))
        command = [
            'inkscape',
            '--pdf-poppler',
            '--export-plain-svg',
            f'--export-filename={svg_path}',
            input_path
        ]

        process = subprocess.run(command,
                               capture_output=True,
                               text=True)
        
        if process.returncode != 0:
            raise Exception(f"Inkscape error: {process.stderr}")
        
        return svg_path
    
    def _convert_with_opencv(self, input_path):
        svg_path = str(Path(input_path).with_suffix('.svg'))
        
        # Read and process image
        img = cv2.imread(input_path)
        if img is None:
            raise ValueError("Failed to load image")
        
        # Convert to grayscale
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # Apply threshold
        _, thresh = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)
        
        # Find contours
        contours, _ = cv2.findContours(thresh, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        
        # Create SVG content
        svg_content = ['<?xml version="1.0" encoding="UTF-8" standalone="no"?>',
                      '<svg xmlns="http://www.w3.org/2000/svg" version="1.1"',
                      f'     width="{img.shape[1]}" height="{img.shape[0]}">']
        
        # Add paths for each contour
        for contour in contours:
            if len(contour) > 2:  # Only process contours with at least 3 points
                path = "M"
                for point in contour:
                    x, y = point[0]
                    path += f" {x},{y}"
                path += " Z"
                svg_content.append(f'    <path d="{path}" style="fill:none;stroke:black;stroke-width:1"/>')
        
        svg_content.append('</svg>')
        
        # Write SVG file
        with open(svg_path, 'w') as f:
            f.write('\n'.join(svg_content))
        
        return svg_path

    

    def generate_gcode(self, svg_path):
        try:
            gcode_path = str(Path(svg_path).with_suffix('.gcode'))
            
            # Read SVG with vpype
            document = vpype.read_svg(svg_path, quantization=0.1)
            
            # Get line data from document (document is a tuple of LineCollections)
            lines = document[0]  # First LineCollection
            
            with open(gcode_path, 'w') as f:
                # Write GCode header
                f.write("; Generated for Automated Writing Machine\n")
                f.write("G21 ; Set units to millimeters\n")
                f.write("G90 ; Absolute positioning\n")
                f.write("G0 Z1.0 F500 ; Move to safe height\n")
                f.write("G28 X Y ; Home X Y\n")
                
                # Process each line in the collection
                for line in lines:
                    # Pen up
                    f.write("G0 Z0.5 F500\n")
                    # Move to start position
                    f.write(f"G0 X{line[0][0]:.3f} Y{line[0][1]:.3f} F3000\n")
                    # Pen down
                    f.write("G1 Z0.0 F200\n")
                    # Draw to end position
                    f.write(f"G1 X{line[1][0]:.3f} Y{line[1][1]:.3f} F2000\n")
                
                # GCode footer
                f.write("G0 Z1.0 F500 ; Lift pen\n")
                f.write("G28 X Y ; Return home\n")
                f.write("M84 ; Disable motors\n")
            
            return gcode_path
        
        except Exception as e:
            self.logger.error(f"Error generating gcode: {str(e)}")
            raise

    def process_file(self, input_path, output_folder="gcode_output"):
        try:
            os.makedirs(output_folder, exist_ok=True)
            file_ext = Path(input_path).suffix.lower()

            # Convert documents to image if needed
            if file_ext in ['.pdf', '.docx', '.doc']:
                # Convert to image
                if file_ext == '.pdf':
                    images = convert_from_path(input_path)
                    temp_img_path = os.path.join(output_folder, "temp_page.png")
                    images[0].save(temp_img_path, 'PNG')
                    input_path = temp_img_path
                elif file_ext in ['.docx', '.doc']:
                    pdf_path = os.path.join(output_folder, "temp.pdf")
                    convert(input_path, pdf_path)
                    # Then convert PDF to image
                    images = convert_from_path(pdf_path)
                    temp_img_path = os.path.join(output_folder, "temp_page.png")
                    images[0].save(temp_img_path, 'PNG')
                    input_path = temp_img_path
            
            # Process image
            img = cv2.imread(input_path)
            if img is None:
                raise ValueError(f"Failed to read image: {input_path}")
            
            # Image processing
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            blurred = cv2.GaussianBlur(gray, (5, 5), 0)
            edges = cv2.Canny(blurred, 50, 150)
            contours, _ = cv2.findContours(edges, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)

            # Generate GCode
            output_path = os.path.join(output_folder, Path(input_path).stem + '.gcode')
            with open(output_path, 'w') as f:
                # Header
                f.write("; Generated for Writing Machine\n")
                f.write("G21 ; Millimeters\n")
                f.write("G90 ; Absolute positioning\n")
                f.write("G28 ; Home all axes\n")

                for contour in contours:
                    if len(contour) > 2:  # Skip tiny contours
                        start = contour[0][0]
                        f.write(f"G0 Z2.0 F1000 ; Pen up\n")
                        f.write(f"G0 X{start[0]} Y{start[1]} F3000\n")
                        f.write(f"G0 Z0.0 F500 ; Pen down\n")
                        
                        for point in contour[1:]:
                            x, y = point[0]
                            f.write(f"G1 X{x} Y{y} F2000\n")

                # Footer
                f.write("G0 Z5.0 F1000 ; Final pen up\n")
                f.write("G28 X Y ; Return home\n")
                f.write("M84 ; Motors off\n")

            # Cleanup temp files
            if file_ext in ['.pdf', '.docx', '.doc']:
                if os.path.exists(temp_img_path):
                    os.remove(temp_img_path)
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)


            return output_path


        except Exception as e:
            self.logger.error(f"Error Processing the file: str{e}")
            raise