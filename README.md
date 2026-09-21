# Benthic Photo Tagger and Organizer

This repository organizes marine survey photos by transect and point, matches them to GPS coordinates, extracts EXIF metadata, and exports a structured Excel workbook with the same column layout and styling as the reference file.

## Features

- Loads coordinate records from CSV files in the coordinates folder. Accept CSV files
- Parses survey metadata such as day, transect, point ID, date, and time. Must be consistent across photos
- Converts latitude/longitude into UTM coordinates with automatic zone detection
- Renames and copies photos into `output/organized_photos/<transect>/`
- Extracts EXIF metadata including camera model, shutter speed, aperture, ISO, and white balance
- Exports a workbook named `output/Survey_Organized.xlsx` can be configured accordingly

## Proposed Repository layout

- `Photos/` — source photos to be organized
- `coords/` — GPS CSV files
- `output/` — generated organized photos and Excel report
- `photo_coordinate_organizer.py` — main processing logic
- `photo_coordinate_organizer.ipynb` — interactive notebook workflow
- `Example.xlsx` — reference workbook layout

## Requirements

- Python 3.10+
- `Pillow`
- `openpyxl`

## Installation

```bash
python -m venv .venv
. .venv/bin/activate  # Windows: .venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -e .
```

## Usage

### CLI

```bash
python photo_coordinate_organizer.py \
  --photos Photos \
  --coords coords \
  --output output \
  --team A \
  --copy-mode copy \
  --utm-zone auto \
  --utm-format MGRS
```

### Python API

```python
from photo_coordinate_organizer import organize_photos

records, excel_file = organize_photos(
    photos_dir="Photos",
    coords_dir="coords",
    output_dir="output",
    team="A",
    copy_mode="copy",
    utm_zone="auto",
    utm_format="MGRS",
)

print(len(records))
print(excel_file)
```

## Notebook workflow

Open `photo_coordinate_organizer.ipynb` and run the cells in order:

1. Setup & imports
2. Inspect coordinates in `coords/`
3. Inspect photos and EXIF metadata
4. Run organization and Excel generation
5. Preview generated table

### Google Colab

```python
!pip install -e .
```

For persistent input and output, mount Google Drive and point the notebook paths
at folders in Drive:

```python
from google.colab import drive
from pathlib import Path

drive.mount('/content/drive')
PHOTOS_DIR = Path('/content/drive/MyDrive/Photos')
COORDS_DIR = Path('/content/drive/MyDrive/coords')
OUTPUT_DIR = Path('/content/drive/MyDrive/output')
```

Use `copy_mode='copy'` in Colab. Hardlinks are not generally supported by
Google Drive-backed filesystems; the code falls back to copying when possible.

## Notes

- The script intentionally leaves the annotation column empty for manual marine interpretation work.
- UTM conversion supports Indonesian survey areas such as Bangka Belitung and Sulawesi.
- The default output uses the MGRS naming convention, such as `X_UTM48M`, `X_UTM51M`.

## License

This project is provided for internal survey and photo-processing workflows. Update the license text to match your deployment or institutional requirements before publishing the repository.
