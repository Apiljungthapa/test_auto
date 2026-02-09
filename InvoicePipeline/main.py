
import sys
from pathlib import Path
from api_client import extract_invoice


def main():
    """Main function to process a PDF file"""
    
    # Check command line arguments
    if len(sys.argv) != 2:
        print("Usage: python main.py <pdf_file_path>")
        print("Example: python main.py 'Docs/invoice.pdf'")
        sys.exit(1)
    
    pdf_path = Path(sys.argv[1])
    
    # Check if file exists
    if not pdf_path.exists():
        print(f"❌ Error: File not found: {pdf_path}")
        sys.exit(1)
    
    # Check if it's a PDF
    if pdf_path.suffix.lower() != '.pdf':
        print(f"⚠️  Warning: File is not a PDF: {pdf_path}")
    
    try:

        json_path = extract_invoice(pdf_path)
        
        print("\n" + "="*50)
        print(f"✅ Processing Complete!")
        print(f"📄 Original PDF: {pdf_path}")
        print(f"📊 JSON Result: {json_path}")
        
        print("\nTo generate reports, run:")
        print(f'python -m reports.lenovo_report "{json_path}"')
        print(f'python -m reports.meyer_report "{json_path}"')
        print("="*50)
        
    except Exception as e:
        print(f"\n❌ Error occurred: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()