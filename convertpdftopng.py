import fitz  # PyMuPDF
import os
 
pdf_path = r"pathgoeshere"
output_folder = r"pathgoeshere"
os.makedirs(output_folder, exist_ok=True)
 
doc = fitz.open(pdf_path)
for page_index in range(len(doc)):
    page = doc[page_index]
    pix = page.get_pixmap(dpi=300)  # High resolution
    img_path = os.path.join(output_folder, f"page{page_index+1}.png")
    pix.save(img_path)
 
print(f"Pages converted to images in: {output_folder}")
