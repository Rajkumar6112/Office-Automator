import typer
from pypdf import PdfWriter
import pandas as pd
from PIL import Image
from pathlib import Path
from reportlab.pdfgen import canvas

app = typer.Typer(help="Office Automator - A small toolkit for daily office tasks.")

# -------- Existing Commands --------
@app.command()
def merge_pdfs(output: str, pdfs: list[str]):
    """Merge multiple PDF files into one."""
    try:
        merger = PdfWriter()
        for pdf in pdfs:
            path = Path(pdf)
            if not path.exists():
                typer.echo(f"❌ File not found: {pdf}")
                return
            merger.append(str(path))
        merger.write(output)
        merger.close()
        typer.echo(f"✅ PDFs merged into {output}")
    except Exception as e:
        typer.echo(f"⚠️ Error: {e}")

@app.command()
def csv_to_excel(csv_file: str, excel_file: str):
    """Convert a CSV file to Excel format."""
    try:
        path = Path(csv_file)
        if not path.exists():
            typer.echo(f"❌ File not found: {csv_file}")
            return
        df = pd.read_csv(path)
        df.to_excel(excel_file, index=False)
        typer.echo(f"✅ CSV converted to Excel: {excel_file}")
    except Exception as e:
        typer.echo(f"⚠️ Error: {e}")

@app.command()
def images_to_pdf(output: str, images: list[str]):
    """Convert images to a single PDF."""
    try:
        pil_images = []
        for img in images:
            path = Path(img)
            if not path.exists():
                typer.echo(f"❌ File not found: {img}")
                return
            image = Image.open(path).convert("RGB")
            pil_images.append(image)
        pil_images[0].save(output, save_all=True, append_images=pil_images[1:])
        typer.echo(f"✅ Images converted to PDF: {output}")
    except Exception as e:
        typer.echo(f"⚠️ Error: {e}")

# -------- New Commands --------
@app.command()
def text_to_pdf(text_file: str, pdf_file: str):
    """Convert a text file to PDF."""
    try:
        path = Path(text_file)
        if not path.exists():
            typer.echo(f"❌ File not found: {text_file}")
            return
        c = canvas.Canvas(pdf_file)
        with open(path, "r", encoding="utf-8") as f:
            text = f.readlines()
        y = 800
        for line in text:
            c.drawString(50, y, line.strip())
            y -= 20
        c.save()
        typer.echo(f"✅ Text file converted to PDF: {pdf_file}")
    except Exception as e:
        typer.echo(f"⚠️ Error: {e}")

@app.command()
def word_count(file: str):
    """Count words in a text or CSV file."""
    try:
        path = Path(file)
        if not path.exists():
            typer.echo(f"❌ File not found: {file}")
            return
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        words = content.split()
        typer.echo(f"📊 Word count: {len(words)} words")
    except Exception as e:
        typer.echo(f"⚠️ Error: {e}")

@app.command()
def rename_files(folder: str, prefix: str):
    """Batch rename files in a folder with a prefix."""
    try:
        path = Path(folder)
        if not path.exists():
            typer.echo(f"❌ Folder not found: {folder}")
            return
        for i, file in enumerate(path.iterdir(), start=1):
            if file.is_file():
                new_name = prefix.format(i=i) + file.suffix
                file.rename(path / new_name)
        typer.echo(f"✅ Files in {folder} renamed with prefix '{prefix}'")
    except Exception as e:
        typer.echo(f"⚠️ Error: {e}")

import zipfile

@app.command()
def compress_files(output: str, files: list[str]):
    """
    Compress multiple files into a single ZIP archive.
    
    Example:
    office compress-files archive.zip file1.txt file2.jpg
    """
    with zipfile.ZipFile(output, "w") as zipf:
        for file in files:
            zipf.write(file)
    print(f"Compressed files into {output}")

import zipfile

@app.command()
def decompress_files(zip_file: str, output_dir: str):
    """Extract all files from a ZIP archive."""
    try:
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            zip_ref.extractall(output_dir)
        typer.echo(f"Extracted files to {output_dir}")
    except FileNotFoundError:
        typer.echo(f"Error: File {zip_file} not found.")
    except zipfile.BadZipFile:
        typer.echo("Error: The file is not a valid ZIP archive.")

from pdf2image import convert_from_path

@app.command()
def pdf_to_images(pdf_file: str, output_prefix: str = typer.Option("page", help="Prefix for output image files")):
    """Convert a PDF into images (one per page)."""
    try:
        images = convert_from_path(pdf_file)
        for i, img in enumerate(images, start=1):
            img_name = f"{output_prefix}_{i}.jpg"
            img.save(img_name, "JPEG")
        typer.echo(f"Converted {pdf_file} into {len(images)} images.")
    except FileNotFoundError:
        typer.echo(f"Error: File {pdf_file} not found.")
    except Exception as e:
        typer.echo(f"Error: {e}")

if __name__ == "__main__":
    app()
