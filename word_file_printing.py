import win32print
import win32api

def print_docx_file(file_path):
    try:
        # Get the default printer name
        printer_name = win32print.GetDefaultPrinter()

        # Prepare the document info structure
        doc_info = ('My Document', None, None)

        # Open the document file in binary read mode
        with open(file_path, 'rb') as f:
            # Create a handle to the printer
            hPrinter = win32print.OpenPrinter(printer_name)

            # Start the document and page
            win32print.StartDocPrinter(hPrinter, 1, doc_info)
            win32print.StartPagePrinter(hPrinter)

            # Send the file to the printer
            win32api.ShellExecute(0, "print", file_path, None, ".", 0)

            # End the document and page
            win32print.EndPagePrinter(hPrinter)
            win32print.EndDocPrinter(hPrinter)

            # Close the printer handle
            win32print.ClosePrinter(hPrinter)

        print(f"Printing {file_path}...")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    file_path = input("Enter the path of the Word document (.docx): ") + f'.docx'
    print_docx_file(file_path)
