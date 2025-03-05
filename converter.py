'''
This script converts PDF or Word documents into a flipbook-style HTML file with page-turning effects.
It uses the pdf2image library to convert PDF pages to images and the turn.js library for the flipbook effect.
The script can be run as a standalone program or imported as a module for use in other applications.

Usage:
1. Run the script and select a PDF or Word document.
2. The script will convert the document into a flipbook-style HTML file.
3. The HTML file can be opened in a web browser to view the flipbook.

Author: Nelson McFadyen
Last Updated: March, 04, 2025
'''

import os
import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2image import convert_from_path

# For DOCX conversion
try:
    from docx2pdf import convert as docx2pdf_convert
except ImportError:
    docx2pdf_convert = None

# Set your poppler bin path here
POPLER_BIN_PATH = r"C:\Users\Nelso\OneDrive\Desktop\flipbookConverter\bin"

def convert_doc_to_pdf(doc_path, pdf_path):
    if docx2pdf_convert is None:
        messagebox.showerror("Error", "docx2pdf module is not installed. Please install it using 'pip install docx2pdf'.")
        return False
    try:
        # Convert DOCX/DOC to PDF; the output PDF will be saved in the same directory as pdf_path.
        docx2pdf_convert(doc_path, os.path.dirname(pdf_path))
        # Sometimes the generated file name is the same as the DOC file but with .pdf extension.
        if not os.path.exists(pdf_path):
            base = os.path.splitext(os.path.basename(doc_path))[0]
            possible_pdf = os.path.join(os.path.dirname(pdf_path), base + ".pdf")
            if os.path.exists(possible_pdf):
                os.rename(possible_pdf, pdf_path)
        return True
    except Exception as e:
        messagebox.showerror("Conversion Error", f"Error converting Word to PDF: {e}")
        return False

import os

def create_html(html_file, image_files, online_host=""):
    """
    Generates the flipbook HTML file.
    
    Args:
      html_file (str): Path to save the generated HTML file.
      image_files (list): List of image file paths.
      online_host (str): Base URL for images when hosted online.
    """
    
    pages_html = ""
    i = 0
    # Loop through the images one by one.
    while i < len(image_files):
        # For the left page:
        if online_host:
            left_image = f'{online_host}/page{i+1}.jpg'
        else:
            left_image = f'images/{os.path.basename(image_files[i])}'
        left_page = f"background-image: url('{left_image}');" if image_files[i] else "background-color: white;"
        pages_html += f"""
        <div class="page" style="{left_page}"></div>
        """
        # For the right page, only if it exists.
        if i + 1 < len(image_files):
            if online_host:
                right_image = f'{online_host}/page{i+2}.jpg'
            else:
                right_image = f'images/{os.path.basename(image_files[i+1])}'
            right_page = f"background-image: url('{right_image}');" if image_files[i+1] else "background-color: white;"
            pages_html += f"""
        <div class="page" style="{right_page}"></div>
        """
        i += 2

    html_content = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Flipbook</title>
  <style>
    /* Prevent scrolling */
    html, body {{
      height: 100vh;
      margin: 0;
      padding: 0;
      overflow: hidden;
      overscroll-behavior: none;
    }}

    /* Entire body with wavy blue pattern background */
    body {{
      background: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1200 120' preserveAspectRatio='none'><path d='M0,0 C300,50 900,0 1200,50 L1200,120 L0,120 Z' fill='%230072E3' opacity='0.5'/><path d='M0,20 C300,70 900,20 1200,70 L1200,120 L0,120 Z' fill='%2300A0FF' opacity='0.5'/></svg>") no-repeat center center;
      background-size: cover;
      font-family: Arial, sans-serif;
    }}

    /* Flipbook container fills space above the fixed footer */
    #flipbook-container {{
      width: 90vw;
      height: calc(100vh - 60px);
      display: flex;
      justify-content: center;
      align-items: center;
      overflow: hidden;
      position: relative;
      text-align: center;
      background: transparent;
      padding-left: 30px;
    }}

    /* The flipbook itself is centered */
    #flipbook {{
      display: inline-block;
      margin: 0 auto;
    }}

    /* Page styling for flipbook images */
    #flipbook .page {{
      background-size: contain;
      background-repeat: no-repeat;
      background-position: center center;
      background-color: white;
      border: 1px solid #999;
      position: absolute;
    }}

    /* Fixed footer controls at the bottom */
    .controls {{
      position: fixed;
      bottom: 0;
      left: 0;
      right: 0;
      height: 60px;
      text-align: center;
      padding: 10px 0;
      background: #0072E3;
      color: #fff;
      z-index: 1000;
    }}

    .controls button {{
      padding: 15px 25px;
      margin: 5px;
      font-size: 18px;
      cursor: pointer;
    }}

    .controls #pageInfo {{
      display: inline-block;
      margin: 0 20px;
      font-size: 20px;
      vertical-align: middle;
    }}

    /* Increase footer button sizes on mobile */
    @media (max-width: 600px) {{
      .controls button {{
        padding: 25px 35px;
        font-size: 24px;
      }}
      .controls {{
        height: 80px;
      }}
    }}
  </style>
  <!-- jQuery and turn.js from CDN -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.0/jquery.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/turn.js/3/turn.min.js"></script>
</head>
<body>
  <div id="flipbook-container">
    <div id="flipbook">
      {pages_html}
    </div>
  </div>
  <div class="controls">
    <button onclick="$('#flipbook').turn('previous')">Previous</button>
    <div id="pageInfo"></div>
    <button onclick="$('#flipbook').turn('next')">Next</button>
  </div>
  <script>
  $(document).ready(function() {{
    var currentDisplay = null;
    var resizeTimer;
    var footerHeight = 60;
    var originalFlipbookHTML = $("#flipbook").html();

    function getFlipbookDimensions() {{
      var width = Math.floor(window.innerWidth * 0.9);
      var availableHeight = window.innerHeight - footerHeight;
      var mode = checkOrientation();
      var height = mode === "double" ? Math.floor(availableHeight * 1.0) : Math.floor(availableHeight * 0.9);
      console.log("Calculated dimensions: width = " + width + ", height = " + height);
      return {{ width: width, height: height }};
    }}

    function checkOrientation() {{
      var mode = window.matchMedia("(orientation: portrait)").matches ? "single" : "double";
      console.log("Orientation check: " + mode);
      return mode;
    }}

    function updatePageInfo() {{
      var total = $("#flipbook").turn("pages");
      var view = $("#flipbook").turn("view");
      var infoText = "";
      if(currentDisplay === "double") {{
        infoText = "Pages: " + view.join(" and ") + " of " + total;
      }} else {{
        infoText = "Page: " + view[0] + " of " + total;
      }}
      $("#pageInfo").text(infoText);
      console.log("Updated page info: " + infoText);
    }}

    function logDebugInfo() {{
      var pagesCount = $("#flipbook").turn("pages");
      var view = $("#flipbook").turn("view");
      console.log("Debug Info:");
      console.log(" - Current display mode: " + currentDisplay);
      console.log(" - Total pages: " + pagesCount);
      console.log(" - Current view (visible pages): ", view);
    }}

    function initFlipbook(displayType) {{
      currentDisplay = displayType;
      var dims = getFlipbookDimensions();
      $("#flipbook").turn({{
        width: dims.width,
        height: dims.height,
        autoCenter: true,
        display: displayType,
        elevation: 50,
        gradients: true,
        when: {{
          turned: function(event, page) {{
            console.log("Turned to page: " + page);
            logDebugInfo();
            updatePageInfo();
          }}
        }}
      }});
      console.log("Initialized flipbook in " + displayType + " mode.");
      logDebugInfo();
      updatePageInfo();
    }}

    function reinitFlipbook(newDisplay) {{
      var currentPage = $("#flipbook").turn("page");
      console.log("Reinitializing flipbook. New display: " + newDisplay + ", current page: " + currentPage);
      try {{
        $("#flipbook").turn("destroy");
        console.log("Destroyed current flipbook instance.");
      }} catch (e) {{
        console.log("Error destroying flipbook instance: " + e);
      }}
      var $oldFlipbook = $("#flipbook");
      $oldFlipbook.replaceWith("<div id='flipbook'>" + originalFlipbookHTML + "</div>");
      setTimeout(function() {{
        initFlipbook(newDisplay);
        var totalPages = $("#flipbook").turn("pages");
        if (currentPage > totalPages) {{
          currentPage = totalPages;
        }}
        $("#flipbook").turn("page", currentPage);
        console.log("Restored to page: " + currentPage);
      }}, 300);
    }}

    initFlipbook(checkOrientation());

    function handleResize() {{
      var newDisplay = checkOrientation();
      console.log("Handle resize: new display: " + newDisplay + ", current display: " + currentDisplay);
      if (newDisplay !== currentDisplay) {{
        reinitFlipbook(newDisplay);
      }} else {{
        var dims = getFlipbookDimensions();
        $("#flipbook").turn("size", dims.width, dims.height);
        updatePageInfo();
      }}
    }}

    $(window).on("resize orientationchange", function() {{
      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(handleResize, 300);
    }});
    
    // Add swipe gesture support to the flipbook container with reduced sensitivity.
    var touchStartX = 0;
    var touchStartY = 0;
    var touchEndX = 0;
    var touchEndY = 0;
    $("#flipbook-container").on("touchstart", function(e) {{
      var touch = e.originalEvent.touches[0];
      touchStartX = touch.clientX;
      touchStartY = touch.clientY;
    }});
    $("#flipbook-container").on("touchend", function(e) {{
      var touch = e.originalEvent.changedTouches[0];
      touchEndX = touch.clientX;
      touchEndY = touch.clientY;
      var diffX = touchEndX - touchStartX;
      var diffY = touchEndY - touchStartY;
      // Only trigger if horizontal movement is at least 75px and more than twice the vertical movement.
      if (Math.abs(diffX) > 75 && Math.abs(diffX) > 2 * Math.abs(diffY)) {{
        // Only trigger if the initial touch was near the left or right edge.
        var containerWidth = $("#flipbook-container").width();
        if (touchStartX < 100 || touchStartX > containerWidth - 100) {{
          if (diffX > 0) {{
            $("#flipbook").turn("previous");
          }} else {{
            $("#flipbook").turn("next");
          }}
        }}
      }}
    }});
  }});
  </script>
</body>
</html>
"""

    with open(html_file, "w", encoding="utf-8") as f:
        f.write(html_content)


def process_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    # Determine if file is PDF or Word.
    if ext == ".pdf":
        pdf_path = file_path
    elif ext in [".doc", ".docx"]:
        pdf_path = file_path + ".pdf"
        success = convert_doc_to_pdf(file_path, pdf_path)
        if not success:
            return
    else:
        messagebox.showerror("Error", "Unsupported file type. Please choose a PDF or Word document.")
        return

    try:
        # Convert PDF pages to images using the specified poppler bin path.
        images = convert_from_path(pdf_path, poppler_path=POPLER_BIN_PATH)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert PDF to images: {e}")
        return

    # Ask user where to save the flipbook.
    out_dir = filedialog.askdirectory(title="Select Output Directory")
    if not out_dir:
        return

    # Create a subfolder for the flipbook.
    flipbook_dir = os.path.join(out_dir, "flipbook")
    os.makedirs(flipbook_dir, exist_ok=True)

    # Create images folder inside the flipbook folder.
    images_dir = os.path.join(flipbook_dir, "images")
    os.makedirs(images_dir, exist_ok=True)

    image_files = []
    for i, image in enumerate(images):
        image_file = os.path.join(images_dir, f"page{i+1}.jpg")
        image.save(image_file, "JPEG")
        # Use a relative path for the HTML.
        image_files.append(f"images/page{i+1}.jpg")

    # Generate the HTML file for the flipbook.
    html_file = os.path.join(flipbook_dir, "index.html")
    create_html(html_file, image_files)
    messagebox.showinfo("Success", f"Flipbook created successfully at:\n{flipbook_dir}")

def main():
    # Create a simple Tkinter UI.
    root = tk.Tk()
    root.title("PDF/Word to Flipbook Converter")
    root.iconbitmap("Assets/projectIcon.ico")
    root.geometry("500x200")

    def browse_file():
        file_path = filedialog.askopenfilename(
            filetypes=[("PDF files", "*.pdf"), ("Word Documents", "*.docx *.doc")],
            title="Select a PDF or Word Document"
        )
        if file_path:
            entry.delete(0, tk.END)
            entry.insert(0, file_path)

    def convert_file():
        file_path = entry.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("Error", "Please select a valid file.")
            return
        process_file(file_path)

    label = tk.Label(root, text="Select PDF or Word Document:")
    label.pack(pady=10)

    entry = tk.Entry(root, width=60)
    entry.pack(pady=5)

    browse_button = tk.Button(root, text="Browse", command=browse_file)
    browse_button.pack(pady=5)

    convert_button = tk.Button(root, text="Convert to Flipbook", command=convert_file)
    convert_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()