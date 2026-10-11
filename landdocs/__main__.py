import argparse
from .engine import generate_docx
from .model import load_client
from .output import generate_client_docx


def main():
    parser = argparse.ArgumentParser(description="Local land-record Word document generator")
    parser.add_argument("--data", help="Client JSON path; omit to open the desktop form")
    parser.add_argument("--output", help="Optional new .docx path; defaults to output/clientname_data.docx")
    args = parser.parse_args()
    if args.data:
        try:
            client = load_client(args.data)
            result = generate_docx(client, args.output) if args.output else generate_client_docx(client)
        except (ValueError, OSError) as error:
            parser.exit(1, f"Error: {error}\n")
        print(f"Generated {result.path}; updated {result.land_tables_updated} land tables.")
    elif args.output:
        parser.error("--output requires --data")
    else:
        try:
            from .gui import run
            run()
        except ImportError as error:
            parser.exit(1, f"Tkinter is unavailable: {error}. Install Python with Tk support; see README.md.\n")
        except RuntimeError as error:
            parser.exit(1, f"Cannot start the desktop application: {error}\n")


if __name__ == "__main__":
    main()
