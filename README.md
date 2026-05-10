```md
# Images to PPT Converter

Convert a sequence of images into a clean PowerPoint presentation automatically using Python.

This project was used to generate a 15-slide presentation from JPEG images for a Computer Organization & Architecture assignment.

---

## Author

- **Name:** M.Hassan.Ali
- **Roll No:** 2024F-BCE-149
- **Department:** Computer Engineering
- **Course:** CE-207T Computer Organization & Architecture

---

## Project Purpose

This script takes multiple image files such as `.jpg`, `.jpeg`, and `.png` and places each image on a separate PowerPoint slide.

It is useful for:

- assignment presentations
- converting screenshots into slides
- creating slide decks from exported HTML images
- quick classroom and projector-ready presentations

---

## Features

- Converts images into a `.pptx` presentation
- Supports:
  - `.jpg`
  - `.jpeg`
  - `.png`
- Detects multiple naming formats:
  - `01.jpg`
  - `1.jpg`
  - `slide-01.jpg`
  - `slide_01.jpg`
- Adds:
  - dark background
  - slide number
  - footer with student and course information
- Creates up to **15 slides**

---

## Project Structure


images-to-ppt-converter/
│
├── images_to_ppt.py
├── README.md
├── .gitignore
│
└── images/
    ├── 01.jpg
    ├── 02.jpg
    ├── 03.jpg
    ├── ...
    └── 15.jpg


---

## Requirements

- Python 3.x
- `python-pptx`

Install dependency:

```bash
pip install python-pptx
```

---

## How to Run

1. Put your images inside the `images/` folder.
2. Make sure they are named in order, for example:


01.jpg
02.jpg
03.jpg
...
15.jpg


3. Run the script:

```bash
python images_to_ppt.py
```

4. The output file will be generated as:


USB_Presentation_M_Hassan_Ali.pptx

---

## Example Use Case

This project was used to:

1. Create an HTML presentation on Universal Serial Bus (USB)
2. Export slides as JPEG images
3. Convert those images into a PowerPoint presentation
4. Present it in class using a projector

---

## Supported Naming Patterns

The script can detect image names like:


01.jpg
1.jpg
slide-01.jpg
slide_01.jpg
Slide01.jpg
Slide1.jpg


---

## Future Improvements

Possible future upgrades:

- automatic slide titles
- custom themes
- better image fit handling
- support for more than 15 slides
- drag-and-drop GUI
- PDF export support

---

## License

This project is for educational and personal use.
```