from getdocs import getDocs
import argparse
import os
import logging

def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('main.log'),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)

def main():
    logger = setup_logging()
    
    parser = argparse.ArgumentParser(description='Automated Writing Machine Controller')
    parser.add_argument('--input', '-i', type=str, required=True, 
                       help='Input file (PDF, DOCX, or image)')
    parser.add_argument('--output', '-o', type=str, default='gcode_output', 
                       help='Output folder for GCode')
    parser.add_argument('--threshold', '-t', type=int, default=127,
                       help='Image threshold (0-255)')
    parser.add_argument('--feed-rate', '-f', type=int, default=2000,
                       help='Drawing speed (mm/min)')
    parser.add_argument('--z-up', '-zu', type=float, default=2.0,
                       help='Pen up height (mm)')
    parser.add_argument('--z-down', '-zd', type=float, default=0.0,
                       help='Pen down height (mm)')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.input):
        logger.error(f"Input file not found: {args.input}")
        return
    
    try:
        logger.info(f"Processing: {args.input}")
        converter = getDocs()
        result = converter.process_file(
            input_path=args.input,
            output_folder=args.output
        )
        logger.info(f"GCode generated: {result}")
        print(f"\nSuccess! Output: {result}")
        
    except Exception as e:
        logger.error(f"Error: {str(e)}")
        raise

if __name__ == '__main__':
    main()