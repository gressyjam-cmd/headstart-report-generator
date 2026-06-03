import os
import json
import datetime
import re
import threading
import tempfile
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from tkinterdnd2 import TkinterDnD
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
from PIL import Image, ImageTk
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT

# Always resolve file paths relative to this script's folder.
# sys._MEIPASS is set by PyInstaller when running as a bundled .exe —
# if it's not set, we're running normally from the source folder.
import sys
if getattr(sys, "frozen", False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Maximum pixel dimension for images embedded in the report.
# Phone photos are often 3000–6000px wide; shrinking them to 1500px
# dramatically reduces file size and generation time with no visible loss.
MAX_IMAGE_PX = 1500

def compress_image_for_report(src_path):
    """
    Return a path to a compressed copy of the image suitable for embedding.
    If the image is already small enough, returns the original path unchanged.
    The compressed copy is written to a temp file and cleaned up automatically
    when the program closes.
    """
    try:
        img = Image.open(src_path)
        # Convert to RGB so JPEG compression works on any mode (RGBA, palette, etc.)
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        w, h = img.size
        if max(w, h) <= MAX_IMAGE_PX:
            return src_path  # already small enough — use as-is
        # Resize, preserving aspect ratio
        img.thumbnail((MAX_IMAGE_PX, MAX_IMAGE_PX), Image.LANCZOS)
        tmp = tempfile.NamedTemporaryFile(suffix=".jpg", delete=False)
        img.save(tmp.name, "JPEG", quality=85, optimize=True)
        tmp.close()
        return tmp.name
    except Exception:
        return src_path  # fall back to original if anything goes wrong

def _compress_paths(paths):
    """Compress a list of image paths, skipping None / missing entries."""
    return [compress_image_for_report(p) if p and os.path.exists(p) else p
            for p in paths]

def _set_img(var_name, value):
    """Safely set an image BooleanVar by name — does nothing if not defined yet."""
    var = globals().get(var_name)
    if isinstance(var, tk.BooleanVar):
        var.set(value)

# practitioner mode
def show_practitioner_frame():
    frame_patient.pack_forget()
    frame_practitioner.pack(fill="both", expand=True)
    bind_mousewheel_to_practitioner()

# patient mode
def show_patient_frame():
    frame_practitioner.pack_forget()
    frame_patient.pack(fill="both", expand=True)
    bind_mousewheel_to_patient()

from docx.shared import Inches

IMAGE_MAP = {
    "right_plagio":           os.path.join(BASE_DIR, "images", "right_plagio.jpg"),
    "left_plagio":            os.path.join(BASE_DIR, "images", "left_plagio.jpg"),
    "temporal_rotation":      os.path.join(BASE_DIR, "images", "temporal_rotation.jpg"),
    "supernumerary_bones_1":  os.path.join(BASE_DIR, "images", "PSCB_1.jpg"),
    "supernumerary_bones_2":  os.path.join(BASE_DIR, "images", "PSCB_2.jpg"),
    "parietal_bone":          os.path.join(BASE_DIR, "images", "parietal_bone.jpg"),
    "sagittal_suture":        os.path.join(BASE_DIR, "images", "sagittal_suture.jpg"),
    "normal_lip_frenulum":    os.path.join(BASE_DIR, "images", "normal_lip_frenulum.jpg"),
    "normal_tongue_frenulum": os.path.join(BASE_DIR, "images", "normal_tongue_frenulum.jpg"),
    "normal_tongue_frenulum_1": os.path.join(BASE_DIR, "images", "normal_tongue_frenulum_1.jpg"),
    "normal_tongue_frenulum_2": os.path.join(BASE_DIR, "images", "normal_tongue_frenulum_2.jpg"),
    "eop":                    os.path.join(BASE_DIR, "images", "EOP.jpg"),
}

# Display titles for each reference image (shown as a caption in the report)
IMAGE_TITLES = {
    "right_plagio":           "Plagiocephaly with Right Lateral Strain",
    "left_plagio":            "Plagiocephaly with Left Lateral Strain",
    "temporal_rotation":      "Temporal Bone",
    "supernumerary_bones_1":  "Potential Supernumerary Cranial Bones Example 1",
    "supernumerary_bones_2":  "Potential Supernumerary Cranial Bones Example 2",
    "parietal_bone":          "Parietal Bone",
    "sagittal_suture":        "Sagittal Suture",
    "normal_lip_frenulum":    "Normal Lip Frenulum",
    "normal_tongue_frenulum": "Normal Tongue Frenulum",
    "normal_tongue_frenulum_1": "Normal Tongue Frenulum Example 1",
    "normal_tongue_frenulum_2": "Normal Tongue Frenulum Example 2",
    "eop":                    "External Occipital Protuberance (EOP)",
}

# Stores patient-specific image paths
patient_images = {
    "anterior_view": None,
    "left_profile": None,
    "right_profile": None,
    "posterior_view": None,
    "superior_view": None,

    "plagio": None,
    "cephalohematoma": None,
    "dolichocephaly": None,
    "brachycephaly": None,
    "torticollis": None,
    "tots": None,
    "metopic_ridge": None,
    "occipital_extension": None,
    "frontal_alignment": None,
    "temporal_alignment": None,
    "retrognathia": None,
    "prognathia": None,
    "lip_frenulum": None,
    "tongue_frenulum": None,
}

patient_image_descriptions = {key: "" for key in patient_images.keys()}

# Create the main application window
root = TkinterDnD.Tk()
root.title("Infant Report Generator")
root.geometry("1200x600")  # width x height

# Scrollable Frame

# Patient Frame
frame_patient = tk.Frame(root)
frame_patient.pack(fill="both", expand=True)

canvas_p = tk.Canvas(frame_patient)
canvas_p.pack(side="left", fill="both", expand=True)

scrollbar_p = ttk.Scrollbar(frame_patient, orient="vertical", command=canvas_p.yview)
scrollbar_p.pack(side="right", fill="y")

# The frame where all widgets go
form_frame = tk.Frame(canvas_p)
canvas_p.create_window((0, 0), window=form_frame, anchor="nw")

canvas_p.configure(yscrollcommand=scrollbar_p.set)
form_frame.bind("<Configure>", lambda e: canvas_p.configure(scrollregion=canvas_p.bbox("all")))

# Practitioner Frame
frame_practitioner = tk.Frame(root)

# Notebook fills the frame directly — no outer canvas needed
notebook = ttk.Notebook(frame_practitioner)
notebook.pack(fill="both", expand=True, padx=10, pady=(10, 0))

# Tab Frames
tab_cranial = tk.Frame(notebook)
tab_oral_neck = tk.Frame(notebook)
tab_other = tk.Frame(notebook)
tab_recommend = tk.Frame(notebook)
tab_images = tk.Frame(notebook)
tab_practitioner = tk.Frame(notebook)

for tab in (tab_cranial, tab_oral_neck, tab_other, tab_recommend, tab_images, tab_practitioner):
    tab.pack(fill="both", expand=True)

#add tabs to notebook
notebook.add(tab_cranial, text="Cranial Findings")
notebook.add(tab_oral_neck, text="Oral & Neck")
notebook.add(tab_other, text="Other Findings")
notebook.add(tab_recommend, text="Recommendations")
notebook.add(tab_images, text="Images")
notebook.add(tab_practitioner, text="Practitioner Notes")

# Maps each tab to its inner scrollable canvas (populated after make_scrollable calls below)
tab_canvas_map = {}

def bind_mousewheel_recursive(widget, callback):
    widget.bind("<MouseWheel>", callback)
    for child in widget.winfo_children():
        bind_mousewheel_recursive(child, callback)


def make_scrollable(parent, margin=20):
    # Canvas
    canvas = tk.Canvas(parent, highlightthickness=0)
    canvas.pack(side="left", fill="both", expand=True)

    # Scrollbar
    scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
    scrollbar.pack(side="right", fill="y")

    canvas.configure(yscrollcommand=scrollbar.set)

    # Inner frame
    inner = tk.Frame(canvas)
    inner_window = canvas.create_window((margin, margin), window=inner, anchor="nw")

    # Resize inner frame when canvas size changes
    def resize_inner(event):
        canvas_width = event.width
        canvas.itemconfig(inner_window, width=canvas_width - margin*2)

    canvas.bind("<Configure>", resize_inner)

    # Update scrollregion
    def update_scrollregion(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    inner.bind("<Configure>", update_scrollregion)

    # Mousewheel support
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # Bind mousewheel to everything inside the scrollable frame
    bind_mousewheel_recursive(inner, _on_mousewheel)
    canvas.bind("<MouseWheel>", _on_mousewheel)


    return inner, canvas

EMU_PER_INCH = 914400

def usable_width_inches(doc):
    s = doc.sections[0]
    usable = s.page_width - s.left_margin - s.right_margin
    return usable / EMU_PER_INCH

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)

def prevent_row_split(row):
    """
    Prevent a table row from splitting across a page break.
    Keeps images and their captions together on the same page.
    """
    from docx.oxml import OxmlElement
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    cantSplit = OxmlElement('w:cantSplit')
    trPr.append(cantSplit)

def replace_placeholder_with_image_row(doc, placeholder, image_paths, max_cols=5):
    """
    Replace a placeholder paragraph with a 1-row table containing up to max_cols images.
    This keeps images on the same line.
    """
    # Filter valid images
    paths = [p for p in image_paths if p and os.path.exists(p)]
    if not paths:
        # If nothing to insert, just remove placeholder
        for p in doc.paragraphs:
            if placeholder in p.text:
                delete_paragraph(p)
                return True
        return False

    cols = min(max_cols, len(paths))
    total_w = usable_width_inches(doc)
    cell_w = total_w / cols
    img_w = max(0.8, cell_w - 0.2)  # leave some padding, but don't make images too small

    # Find placeholder paragraph in body
    for p in doc.paragraphs:
        if placeholder in p.text:
            # Create table at end (python-docx limitation), then move it into place
            table = doc.add_table(rows=1, cols=cols)
            table.autofit = False
            table.alignment = WD_TABLE_ALIGNMENT.CENTER

            for i in range(cols):
                cell = table.rows[0].cells[i]
                cell.width = Inches(cell_w)  # Word respects cell widths more reliably than column widths [1](https://pytutorial.com/python-docx-paragraph-formatting-guide/)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

                # Each cell already has an empty paragraph; use it instead of adding one [2](https://skelmis-docx.readthedocs.io/en/stable/user/styles-using.html)
                cp = cell.paragraphs[0]
                cp.text = ""
                cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = cp.add_run()
                run.add_picture(paths[i], width=Inches(img_w))  # place picture precisely here [3](https://stackoverflow.com/questions/54945258/python-docx-unable-to-use-existing-document-no-style-with-name-title)

            # Move table XML directly after the placeholder paragraph
            p._p.addnext(table._tbl)
            delete_paragraph(p)
            return True

    return False




cranial_frame, cranial_canvas = make_scrollable(tab_cranial)
oral_neck_frame, oral_neck_canvas = make_scrollable(tab_oral_neck)
other_frame, other_canvas = make_scrollable(tab_other)
recommend_frame, recommend_canvas = make_scrollable(tab_recommend)
practitioner_frame, practitioner_canvas = make_scrollable(tab_practitioner)
images_frame, images_canvas = make_scrollable(tab_images)

tab_canvas_map.update({
    str(tab_cranial): cranial_canvas,
    str(tab_oral_neck): oral_neck_canvas,
    str(tab_other): other_canvas,
    str(tab_recommend): recommend_canvas,
    str(tab_images): images_canvas,
    str(tab_practitioner): practitioner_canvas,
})


##=====================
# Recommendations Tab
##=====================

# -----------------------------
# Thrush Recommendation
# -----------------------------

tk.Label(recommend_frame, text="Thrush Recommendation:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20,0))

thrush_rec_box = tk.Text(recommend_frame, width=80, height=4, wrap="word")
thrush_rec_box.pack(anchor="w", padx=20, pady=5)
thrush_rec_box.config(state="disabled")

# -----------------------------
# Cradle Cap Recommendation
# -----------------------------

tk.Label(recommend_frame, text="Cradle Cap Recommendation:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20,0))

cradle_cap_rec_box = tk.Text(recommend_frame, width=80, height=4, wrap="word")
cradle_cap_rec_box.pack(anchor="w", padx=20, pady=5)
cradle_cap_rec_box.config(state="disabled")

# -----------------------------
# Cephalohematoma Recommendation
# -----------------------------

tk.Label(recommend_frame, text="Cephalohematoma Recommendation:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20,0))

cephalohematoma_rec_box = tk.Text(recommend_frame, width=80, height=3, wrap="word")
cephalohematoma_rec_box.pack(anchor="w", padx=20, pady=5)
cephalohematoma_rec_box.config(state="disabled")

# -----------------------------
# Blocked Tear Duct Recommendation
# -----------------------------

tk.Label(recommend_frame, text="Blocked Tear Duct Recommendation:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20,0))

blocked_tear_duct_rec_box = tk.Text(recommend_frame, width=80, height=3, wrap="word")
blocked_tear_duct_rec_box.pack(anchor="w", padx=20, pady=5)
blocked_tear_duct_rec_box.config(state="disabled")

# -----------------------------
# Jaw Exercise Recommendation
# -----------------------------

tk.Label(recommend_frame, text="Jaw Exercise Recommendation:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20,0))

jaw_ex_box = tk.Text(recommend_frame, width=80, height=3, wrap="word")
jaw_ex_box.pack(anchor="w", padx=20, pady=5)
jaw_ex_box.config(state="disabled")

# -----------------------------
# Homeopathy Recommendation
# -----------------------------

homeo_rec = tk.BooleanVar()

def toggle_homeopathy():
    homeo_rec_box.config(state="normal")
    homeo_rec_box.delete("1.0", "end")

    if homeo_rec.get():
        homeo_rec_box.insert("1.0", "Homeopathic remedies have been recommended.")
    else:
        homeo_rec_box.config(state="disabled")

tk.Checkbutton(
    recommend_frame,
    text="Include Homeopathy Recommendation",
    variable=homeo_rec,
    command=lambda: toggle_homeopathy()
).pack(anchor="w", padx=20, pady=(10,0))

tk.Label(recommend_frame, text="Homeopathy Recommendation:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20,0))

homeo_rec_box = tk.Text(recommend_frame, width=80, height=3, wrap="word")
homeo_rec_box.pack(anchor="w", padx=20, pady=5)
homeo_rec_box.config(state="disabled")

# -----------------------------
# Homeopathy Consultation Recommendation
# -----------------------------

homeo_cons_rec = tk.BooleanVar()

def toggle_homeo_consult():
    homeo_cons_rec_box.config(state="normal")
    homeo_cons_rec_box.delete("1.0", "end")

    if homeo_cons_rec.get():
        homeo_cons_rec_box.insert(
            "1.0",
            "A homeopathic consultation with Stephanie Kononovich has been recommended."
        )
    else:
        homeo_cons_rec_box.config(state="disabled")

tk.Checkbutton(
    recommend_frame,
    text="Recommend Homeopathy Consultation",
    variable=homeo_cons_rec,
    command=lambda: toggle_homeo_consult()
).pack(anchor="w", padx=20, pady=(10,0))

tk.Label(recommend_frame, text="Homeopathy Consultation:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20,0))

homeo_cons_rec_box = tk.Text(recommend_frame, width=80, height=3, wrap="word")
homeo_cons_rec_box.pack(anchor="w", padx=20, pady=5)
homeo_cons_rec_box.config(state="disabled")

# -----------------------------
# Primitive Reflex Integration Recommendation
# -----------------------------

primitive_reflex_rec = tk.BooleanVar()

def toggle_primitive_reflex():
    primitive_reflex_rec_box.config(state="normal")
    primitive_reflex_rec_box.delete("1.0", "end")

    if primitive_reflex_rec.get():
        primitive_reflex_rec_box.insert(
            "1.0",
            "Primitive reflex integration consultation with Stephanie Kononovich has been recommended."
        )
    else:
        primitive_reflex_rec_box.config(state="disabled")

tk.Checkbutton(
    recommend_frame,
    text="Recommend Primitive Reflex Integration",
    variable=primitive_reflex_rec,
    command=lambda: toggle_primitive_reflex()
).pack(anchor="w", padx=20, pady=(10,0))

tk.Label(recommend_frame, text="Primitive Reflex Integration:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20,0))

primitive_reflex_rec_box = tk.Text(recommend_frame, width=80, height=4, wrap="word")
primitive_reflex_rec_box.pack(anchor="w", padx=20, pady=5)
primitive_reflex_rec_box.config(state="disabled")

tk.Label(recommend_frame, text="Observed Primitive Reflexes:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20,0))

primitive_reflexes_findings = tk.Text(recommend_frame, width=80, height=4, wrap="word")
primitive_reflexes_findings.pack(anchor="w", padx=20, pady=5)
primitive_reflexes_findings.config(state="normal")

# -----------------------------
# TOTs Dentist List (Multi-line)
# -----------------------------

tots_docs_rec_text = """The following are exceptional, qualified pediatric dentists for an honest evaluation of TOTs:

• Dr. Alina Garciamendez – https://kidspediatricdentistry.com/dr-alina-garciamendez/ (Allen)
• Dr. Matthew Schwed – https://schwedkidsdental.com/ (Garland)
• Dr. Stacy Coe – https://clearforkpediatricdentistry.com/frenectomy/ (Ft. Worth)
• Dr. Melanie Throne – https://www.mvcdds.com/ (Ft. Worth)
• Dr. Blair Goodall – https://www.morethansmilespediatricdentistry.com/ (Frisco)
• Dr. My Matthews – https://www.blossompediatricdentistry.com/ (Prosper)
"""

tk.Label(recommend_frame, text="TOTs Dentist List:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20,0))

tots_docs_rec_box = tk.Text(recommend_frame, width=80, height=10, wrap="word")
tots_docs_rec_box.pack(anchor="w", padx=20, pady=5)
tots_docs_rec_box.insert("1.0", "")
tots_docs_rec_box.config(state="disabled")

# -----------------------------
# Treatment Plan
# -----------------------------

# Weeks at Birth
tk.Label(recommend_frame, text="Treatment Plan:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
treatment_plan = tk.StringVar()
tk.Entry(recommend_frame, textvariable=treatment_plan, width=40).pack(anchor="w", padx=20)



##=====================
# Cranial Findings Tab
##=====================

tk.Label(form_frame, text="Date of Service:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))

date_of_service = tk.StringVar()
tk.Entry(form_frame, textvariable=date_of_service, width=40).pack(anchor="w", padx=20)

# -----------------------------
# Cephalohematoma (Toggle Group)
# -----------------------------

def toggle_cephalohematoma_fields():
    state = "normal" if cephalohematoma.get() else "disabled"
    cephalohematoma_measurement_entry.config(state=state)
    cephalohematoma_location_entry.config(state=state)
    cephalohematoma_cross_suture_dropdown.config(state=state)
    cephalohematoma_calcium_dropdown.config(state=state)
    cephalohematoma_calcium_location_entry.config(state=state)

    # Auto-select reference diagrams
    _set_img("img_parietal_bone", cephalohematoma.get())
    _set_img("img_sagittal_suture", cephalohematoma.get())

    cephalohematoma_rec_box.config(state="normal")
    cephalohematoma_rec_box.delete("1.0", "end")

    if cephalohematoma.get():
        cephalohematoma_rec_box.insert(
            "1.0",
            "Silicea 6x and Calc Fluor 6x 2–3 times daily."
        )
    else:
        cephalohematoma_rec_box.config(state="disabled")



cephalohematoma = tk.BooleanVar()
tk.Checkbutton(
    cranial_frame,
    text="Cephalohematoma",
    variable=cephalohematoma,
    command=toggle_cephalohematoma_fields
).pack(anchor="w", padx=10, pady=(10,0))

# Measurement
tk.Label(cranial_frame, text="Measurement:", font=("Arial", 10)).pack(anchor="w", padx=30)
cephalohematoma_measurement = tk.StringVar()
cephalohematoma_measurement_entry = tk.Entry(cranial_frame, textvariable=cephalohematoma_measurement, width=30)
cephalohematoma_measurement_entry.pack(anchor="w", padx=30)

# Location
tk.Label(cranial_frame, text="Location:", font=("Arial", 10)).pack(anchor="w", padx=30)
cephalohematoma_location = tk.StringVar()
cephalohematoma_location_entry = tk.Entry(cranial_frame, textvariable=cephalohematoma_location, width=30)
cephalohematoma_location_entry.pack(anchor="w", padx=30)

# Crosses sutures?
tk.Label(cranial_frame, text="Crosses sutures?", font=("Arial", 10)).pack(anchor="w", padx=30, pady=(20,0))
cephalohematoma_cross_suture = tk.StringVar()
cephalohematoma_cross_suture_dropdown = ttk.Combobox(
    cranial_frame,
    textvariable=cephalohematoma_cross_suture,
    values=["Choose an Option", "Yes", "No"],
    width=27
)
cephalohematoma_cross_suture_dropdown.pack(anchor="w", padx=30, pady=10)
cephalohematoma_cross_suture_dropdown.current(0)

# Calcium deposits?
tk.Label(cranial_frame, text="Calcium deposits present?", font=("Arial", 10)).pack(anchor="w", padx=30)
cephalohematoma_calcium = tk.StringVar()
cephalohematoma_calcium_dropdown = ttk.Combobox(
    cranial_frame,
    textvariable=cephalohematoma_calcium,
    values=["Choose an Option", "Yes", "No"],
    width=27
)
cephalohematoma_calcium_dropdown.pack(anchor="w", padx=30, pady=10)
cephalohematoma_calcium_dropdown.current(0)

# Calcium deposit location
tk.Label(cranial_frame, text="Calcium deposit location:", font=("Arial", 10)).pack(anchor="w", padx=30)
cephalohematoma_calcium_location = tk.StringVar()
cephalohematoma_calcium_location_entry = tk.Entry(cranial_frame, textvariable=cephalohematoma_calcium_location, width=30)
cephalohematoma_calcium_location_entry.pack(anchor="w", padx=30)

# Disable all fields initially
toggle_cephalohematoma_fields()

# -----------------------------
# Brachycephaly (Toggle Group)
# -----------------------------

def toggle_brachycephaly_fields():
    state = "normal" if brachycephaly.get() else "disabled"
    brachycephaly_severity_dropdown.config(state=state)
    brachycephaly_measurement_entry.config(state=state)

brachycephaly = tk.BooleanVar()
tk.Checkbutton(
    cranial_frame,
    text="Brachycephaly",
    variable=brachycephaly,
    command=toggle_brachycephaly_fields
).pack(anchor="w", padx=10, pady=(20,0))

# Severity
tk.Label(cranial_frame, text="Severity:", font=("Arial", 10)).pack(anchor="w", padx=30)
brachycephaly_severity = tk.StringVar()
brachycephaly_severity_dropdown = ttk.Combobox(
    cranial_frame,
    textvariable=brachycephaly_severity,
    values=["Choose an Option", "Mild", "Moderate", "Severe"],
    width=27
)
brachycephaly_severity_dropdown.pack(anchor="w", padx=30, pady=10)
brachycephaly_severity_dropdown.current(0)

# CI Measurement
tk.Label(cranial_frame, text="Cephalic Index (CI):", font=("Arial", 10)).pack(anchor="w", padx=30)
brachycephaly_measurement = tk.StringVar()
brachycephaly_measurement_entry = tk.Entry(cranial_frame, textvariable=brachycephaly_measurement, width=30)
brachycephaly_measurement_entry.pack(anchor="w", padx=30, pady=10)

# Disable initially
toggle_brachycephaly_fields()

# -----------------------------
# Plagiocephaly (Toggle Group)
# -----------------------------

def update_plagio_image(*_):
    """Select the correct plagiocephaly diagram based on laterality, or clear both."""
    if not plagiocephaly.get():
        _set_img("img_right_plagio", False)
        _set_img("img_left_plagio", False)
        return
    lat = plagiocephaly_laterality.get()
    _set_img("img_right_plagio", lat == "Right")
    _set_img("img_left_plagio", lat == "Left")

def toggle_plagiocephaly_fields():
    state = "normal" if plagiocephaly.get() else "disabled"
    plagiocephaly_severity_dropdown.config(state=state)
    plagiocephaly_measurement_entry.config(state=state)
    plagiocephaly_laterality_dropdown.config(state=state)
    update_plagio_image()

plagiocephaly = tk.BooleanVar()
tk.Checkbutton(
    cranial_frame,
    text="Plagiocephaly",
    variable=plagiocephaly,
    command=toggle_plagiocephaly_fields
).pack(anchor="w", padx=10, pady=(20,0))

# Severity
tk.Label(cranial_frame, text="Severity:", font=("Arial", 10)).pack(anchor="w", padx=30)
plagiocephaly_severity = tk.StringVar()
plagiocephaly_severity_dropdown = ttk.Combobox(
    cranial_frame,
    textvariable=plagiocephaly_severity,
    values=["Choose an Option", "Mild", "Moderate", "Severe"],
    width=27
)
plagiocephaly_severity_dropdown.pack(anchor="w", padx=30, pady=10)
plagiocephaly_severity_dropdown.current(0)

# CVA Measurement
tk.Label(cranial_frame, text="Cranial Vault Asymmetry (CVA):", font=("Arial", 10)).pack(anchor="w", padx=30)
plagiocephaly_measurement = tk.StringVar()
plagiocephaly_measurement_entry = tk.Entry(cranial_frame, textvariable=plagiocephaly_measurement, width=30)
plagiocephaly_measurement_entry.pack(anchor="w", padx=30, pady=10)

# Laterality
tk.Label(cranial_frame, text="Laterality:", font=("Arial", 10)).pack(anchor="w", padx=30)
plagiocephaly_laterality = tk.StringVar()
plagiocephaly_laterality_dropdown = ttk.Combobox(
    cranial_frame,
    textvariable=plagiocephaly_laterality,
    values=["Choose an Option", "Right", "Left"],
    width=27
)
plagiocephaly_laterality_dropdown.pack(anchor="w", padx=30, pady=10)
plagiocephaly_laterality_dropdown.current(0)
plagiocephaly_laterality_dropdown.bind("<<ComboboxSelected>>", update_plagio_image)

# Disable initially
toggle_plagiocephaly_fields()

# -----------------------------
# Dolichocephaly (Toggle Group)
# -----------------------------

def toggle_dolichocephaly_fields():
    state = "normal" if dolichocephaly.get() else "disabled"
    dolichocephaly_severity_dropdown.config(state=state)
    dolichocephaly_measurement_entry.config(state=state)

dolichocephaly = tk.BooleanVar()
tk.Checkbutton(
    cranial_frame,
    text="Dolichocephaly",
    variable=dolichocephaly,
    command=toggle_dolichocephaly_fields
).pack(anchor="w", padx=10, pady=(20,0))

# Severity
tk.Label(cranial_frame, text="Severity:", font=("Arial", 10)).pack(anchor="w", padx=30)
dolichocephaly_severity = tk.StringVar()
dolichocephaly_severity_dropdown = ttk.Combobox(
    cranial_frame,
    textvariable=dolichocephaly_severity,
    values=["Choose an Option", "Mild", "Moderate", "Severe"],
    width=27
)
dolichocephaly_severity_dropdown.pack(anchor="w", padx=30, pady=10)
dolichocephaly_severity_dropdown.current(0)

# CI Measurement
tk.Label(cranial_frame, text="Cephalic Index (CI):", font=("Arial", 10)).pack(anchor="w", padx=30)
dolichocephaly_measurement = tk.StringVar()
dolichocephaly_measurement_entry = tk.Entry(cranial_frame, textvariable=dolichocephaly_measurement, width=30)
dolichocephaly_measurement_entry.pack(anchor="w", padx=30, pady=10)

# Disable initially
toggle_dolichocephaly_fields()

# -----------------------------
# Facial Asymmetry (Toggle Group)
# -----------------------------

facial_asymmetry = tk.BooleanVar()
tk.Checkbutton(
    cranial_frame,
    text="Facial Asymmetry",
    variable=facial_asymmetry
).pack(anchor="w", padx=10, pady=(20,0))

# -----------------------------
# Occipital Extension (Toggle Group)
# -----------------------------

occipital_extension = tk.BooleanVar()
tk.Checkbutton(
    cranial_frame,
    text="Occipital Extension",
    variable=occipital_extension
).pack(anchor="w", padx=10, pady=(20,0))

# -----------------------------
# Sutural Restriction (Toggle Group)
# -----------------------------

sutural_restriction = tk.BooleanVar()
tk.Checkbutton(
    cranial_frame,
    text="Sutural Restriction",
    variable=sutural_restriction
).pack(anchor="w", padx=10, pady=(20,0))

# -----------------------------
# Metopic Ridge (Toggle Group)
# -----------------------------

metopic_ridge = tk.BooleanVar()
tk.Checkbutton(
    cranial_frame,
    text="Metopic Ridge",
    variable=metopic_ridge
).pack(anchor="w", padx=10, pady=(20,0))

# -----------------------------
# supernumerary cranial bones
# -----------------------------

def toggle_supernumerary_bones():
    _set_img("img_pscb_1", supernumerary_bones.get())
    _set_img("img_pscb_2", supernumerary_bones.get())

supernumerary_bones = tk.BooleanVar()
tk.Checkbutton(
    cranial_frame,
    text="Supernumerary Cranial Bones",
    variable=supernumerary_bones,
    command=toggle_supernumerary_bones
).pack(anchor="w", padx=10, pady=(20,0))

# -----------------------------
# Frontal Alignment (Toggle Group)
# -----------------------------

frontal_alignment = tk.BooleanVar()
tk.Checkbutton(
    cranial_frame,
    text="Frontal Alignment",
    variable=frontal_alignment
).pack(anchor="w", padx=10, pady=(20,0))

# -----------------------------
# Temporal Alignment (Toggle Group)
# -----------------------------

def toggle_temporal_alignment_fields():
    state = "normal" if temporal_alignment.get() else "disabled"
    temporal_ear_flare_laterality_dropdown.config(state=state)
    temporal_alignment_outcome_entry.config(state=state)

    # Auto-select reference diagram
    _set_img("img_temporal_rotation", temporal_alignment.get())

temporal_alignment = tk.BooleanVar()
tk.Checkbutton(
    cranial_frame,
    text="Temporal Alignment",
    variable=temporal_alignment,
    command=toggle_temporal_alignment_fields
).pack(anchor="w", padx=10, pady=(20,0))

# Laterality
tk.Label(cranial_frame, text="Ear Flare Laterality:", font=("Arial", 10)).pack(anchor="w", padx=30)
temporal_ear_flare_laterality = tk.StringVar()
temporal_ear_flare_laterality_dropdown = ttk.Combobox(
    cranial_frame,
    textvariable=temporal_ear_flare_laterality,
    values=["Choose an Option", "Right", "Left"],
    width=27
)
temporal_ear_flare_laterality_dropdown.pack(anchor="w", padx=30, pady=10)
temporal_ear_flare_laterality_dropdown.current(0)

# Outcome
tk.Label(cranial_frame, text="Outcome:", font=("Arial", 10)).pack(anchor="w", padx=30)
temporal_alignment_outcome = tk.StringVar()
temporal_alignment_outcome_entry = tk.Entry(cranial_frame, textvariable=temporal_alignment_outcome, width=30)
temporal_alignment_outcome_entry.pack(anchor="w", padx=30, pady=10)

toggle_temporal_alignment_fields()

##=====================
# Oral & Neck Findings
##=====================

# -----------------------------
# Torticollis (Toggle Group)
# -----------------------------

def toggle_torticollis_fields():
    state = "normal" if torticollis.get() else "disabled"
    torticollis_laterality_dropdown.config(state=state)

torticollis = tk.BooleanVar()
tk.Checkbutton(
    oral_neck_frame,
    text="Torticollis",
    variable=torticollis,
    command=toggle_torticollis_fields
).pack(anchor="w", padx=10, pady=(10,0))

# Laterality
tk.Label(oral_neck_frame, text="Laterality:", font=("Arial", 10)).pack(anchor="w", padx=30)
torticollis_laterality = tk.StringVar()
torticollis_laterality_dropdown = ttk.Combobox(
    oral_neck_frame,
    textvariable=torticollis_laterality,
    values=["Choose an Option", "Right", "Left"],
    width=27
)
torticollis_laterality_dropdown.pack(anchor="w", padx=30, pady=10)
torticollis_laterality_dropdown.current(0)

toggle_torticollis_fields()

# -----------------------------
# Tethered Oral Tissues (TOTs)
# -----------------------------

def toggle_tethered_oral_tissues_fields():
    if tethered_oral_tissues.get():
        tots_docs_rec_box.config(state="normal")
        tots_docs_rec_box.delete("1.0", "end")
        tots_docs_rec_box.insert("1.0", tots_docs_rec_text)
    else:
        tots_docs_rec_box.config(state="normal")
        tots_docs_rec_box.delete("1.0", "end")
        tots_docs_rec_box.config(state="disabled")

    # Auto-select reference diagrams
    _set_img("img_normal_lip_frenulum", tethered_oral_tissues.get())
    _set_img("img_normal_tongue_frenulum_1", tethered_oral_tissues.get())
    _set_img("img_normal_tongue_frenulum_2", tethered_oral_tissues.get())


tethered_oral_tissues = tk.BooleanVar()
tk.Checkbutton(
    oral_neck_frame,
    text="Tethered Oral Tissues (TOTs)",
    variable=tethered_oral_tissues,
    command=toggle_tethered_oral_tissues_fields
).pack(anchor="w", padx=10, pady=(20,0))


# -----------------------------
# Internal Maxilla Rotation
# -----------------------------

internal_maxilla_rotation = tk.BooleanVar()
tk.Checkbutton(
    oral_neck_frame,
    text="Internal Maxilla Rotation",
    variable=internal_maxilla_rotation
).pack(anchor="w", padx=10, pady=(20,0))

# -----------------------------
# Retrognathia/Prognathia rec bind
# -----------------------------

def update_jaw_exercises():
    jaw_ex_box.config(state="normal")
    jaw_ex_box.delete("1.0", "end")

    if retrognathia.get() or prognathia.get():
        jaw_ex_box.insert("1.0", "Jaw exercises have been recommended.")
    else:
        jaw_ex_box.config(state="disabled")


# -----------------------------
# Retrognathia
# -----------------------------

def toggle_retrognathia_fields():
    jaw_ex_box.config(state="normal")
    jaw_ex_box.delete("1.0", "end")

    if retrognathia.get():
        jaw_ex_box.insert("1.0", "Jaw exercises have been recommended.")
    else:
        jaw_ex_box.config(state="disabled")


retrognathia = tk.BooleanVar()
tk.Checkbutton(
    oral_neck_frame,
    text="Retrognathia",
    variable=retrognathia,
    command=update_jaw_exercises
).pack(anchor="w", padx=10, pady=(20,0))



# -----------------------------
# Prognathia
# -----------------------------

prognathia = tk.BooleanVar()
tk.Checkbutton(
    oral_neck_frame,
    text="Prognathia",
    variable=prognathia,
    command=update_jaw_exercises
).pack(anchor="w", padx=10, pady=(20,10))



##=====================
# Other Findings Tab
##=====================

# -----------------------------
# Thrush (Toggle Group)
# -----------------------------

def toggle_thrush_fields():
    thrush_rec_box.config(state="normal")
    thrush_rec_box.delete("1.0", "end")

    if thrush.get():
        thrush_rec_box.insert(
            "1.0",
            "Lacto Prime or Baby Formula Probiotic by SFI\n"
            "Anovite Colostrum for baby — start with 1/2 tsp daily, increase to 1 Tbsp daily\n"
            "Appointment with Stephanie Kononovich https://foundationhomeopathy.com\n"
            "Ther-biotic Complete (1/4 tsp) for mom twice daily (if breastfeeding)\n"
            "Mom eliminate sugar and simple carbs (if breastfeeding)\n"
            "Colloidal Silver (Argentyn 23) sprayed on nipples before and after nursing (if breastfeeding)\n"
            "Keep nipples aired out — moist and dark make yeast grow (if breastfeeding)\n"
            "Silverette Nipple Shields (if breastfeeding)"
        )
    else:
        thrush_rec_box.config(state="disabled")


thrush = tk.BooleanVar()
tk.Checkbutton(
    other_frame,
    text="Thrush",
    variable=thrush,
    command=toggle_thrush_fields
).pack(anchor="w", padx=10, pady=(10,0))

toggle_thrush_fields()


# -----------------------------
# Cradle Cap (Toggle Group)
# -----------------------------

def toggle_cradle_cap_fields():
    cradle_cap_rec_box.config(state="normal")
    cradle_cap_rec_box.delete("1.0", "end")

    if cradle_cap.get():
        cradle_cap_rec_box.insert(
            "1.0",
            "Calc Sulph 6x 3x/day\n"
            "Lacto Prima or Baby Formula Probiotic by SFI\n"
            "Argentyn 23 silver gel or spray on area 3x/day, then massage coconut or olive oil into scalp (do not scrub)"
        )
    else:
        cradle_cap_rec_box.config(state="disabled")



cradle_cap = tk.BooleanVar()
tk.Checkbutton(
    other_frame,
    text="Cradle Cap",
    variable=cradle_cap,
    command=toggle_cradle_cap_fields
).pack(anchor="w", padx=10, pady=(20,0))

toggle_cradle_cap_fields()


# -----------------------------
# Blocked Tear Duct (Toggle Group)
# -----------------------------

def toggle_blocked_tear_duct_fields():
    state = "normal" if blocked_tear_duct.get() else "disabled"
    blocked_tear_duct_laterality_dropdown.config(state=state)

    blocked_tear_duct_rec_box.config(state="normal")
    blocked_tear_duct_rec_box.delete("1.0", "end")

    if blocked_tear_duct.get():
        blocked_tear_duct_rec_box.insert(
            "1.0",
            "Silicea 6x 2–3x daily\n"
            "Argentyn 23 silver (spray or drops) in the eye 2x daily"
        )
    else:
        blocked_tear_duct_rec_box.config(state="disabled")



blocked_tear_duct = tk.BooleanVar()
tk.Checkbutton(
    other_frame,
    text="Blocked Tear Duct",
    variable=blocked_tear_duct,
    command=toggle_blocked_tear_duct_fields
).pack(anchor="w", padx=10, pady=(20,0))

# Laterality
tk.Label(other_frame, text="Laterality:", font=("Arial", 10)).pack(anchor="w", padx=30)
blocked_tear_duct_laterality = tk.StringVar()
blocked_tear_duct_laterality_dropdown = ttk.Combobox(
    other_frame,
    textvariable=blocked_tear_duct_laterality,
    values=["Choose an Option", "Right", "Left", "Bilateral"],
    width=27
)
blocked_tear_duct_laterality_dropdown.pack(anchor="w", padx=30, pady=10)
blocked_tear_duct_laterality_dropdown.current(0)

toggle_blocked_tear_duct_fields()

# -----------------------------
# Acetabular Click (Toggle Group)
# -----------------------------

def toggle_acetabular_click_fields():
    state = "normal" if acetabular_click.get() else "disabled"
    acetabular_click_laterality_dropdown.config(state=state)

acetabular_click = tk.BooleanVar()
tk.Checkbutton(
    other_frame,
    text="Acetabular Click",
    variable=acetabular_click,
    command=toggle_acetabular_click_fields
).pack(anchor="w", padx=10, pady=(20,0))

# Laterality
tk.Label(other_frame, text="Laterality:", font=("Arial", 10)).pack(anchor="w", padx=30)
acetabular_click_laterality = tk.StringVar()
acetabular_click_laterality_dropdown = ttk.Combobox(
    other_frame,
    textvariable=acetabular_click_laterality,
    values=["Choose an Option", "Right", "Left", "Bilateral"],
    width=27
)
acetabular_click_laterality_dropdown.pack(anchor="w", padx=30, pady=10)
acetabular_click_laterality_dropdown.current(0)

toggle_acetabular_click_fields()

# -----------------------------
# Femur Rotation (Toggle Group)
# -----------------------------

def toggle_femur_rotation_fields():
    state = "normal" if femur_rotation.get() else "disabled"
    femur_rotation_laterality_dropdown.config(state=state)
    femur_rotation_restriction_dropdown.config(state=state)

femur_rotation = tk.BooleanVar()
tk.Checkbutton(
    other_frame,
    text="Femur Rotation",
    variable=femur_rotation,
    command=toggle_femur_rotation_fields
).pack(anchor="w", padx=10, pady=(20,0))

# Laterality
tk.Label(other_frame, text="Laterality:", font=("Arial", 10)).pack(anchor="w", padx=30)
femur_rotation_laterality = tk.StringVar()
femur_rotation_laterality_dropdown = ttk.Combobox(
    other_frame,
    textvariable=femur_rotation_laterality,
    values=["Choose an Option", "Right", "Left", "Bilateral"],
    width=27
)
femur_rotation_laterality_dropdown.pack(anchor="w", padx=30, pady=10)
femur_rotation_laterality_dropdown.current(0)

# Restriction
tk.Label(other_frame, text="Restriction (internal/external):", font=("Arial", 10)).pack(anchor="w", padx=30)
femur_rotation_restriction = tk.StringVar()
femur_rotation_restriction_dropdown = ttk.Combobox(
    other_frame,
    textvariable=femur_rotation_restriction,
    values=["Choose an Option", "Internal", "External"],
    width=27
)
femur_rotation_restriction_dropdown.pack(anchor="w", padx=30, pady=10)
femur_rotation_restriction_dropdown.current(0)

toggle_femur_rotation_fields()

# -----------------------------
# Inappropriate Reflexes (Toggle Group)
# -----------------------------

inappropriate_reflexes = tk.BooleanVar()
tk.Checkbutton(
    other_frame,
    text="Inappropriate Reflexes Observed",
    variable=inappropriate_reflexes
).pack(anchor="w", padx=10, pady=(20,0))

# -----------------------------
# Skeletal Alignment
# -----------------------------

skeletal_alignment = tk.BooleanVar()
tk.Checkbutton(
    other_frame,
    text="Skeletal Alignment Findings",
    variable=skeletal_alignment
).pack(anchor="w", padx=10, pady=(20,0))


##=====================
# Images Tab
##=====================


def insert_image(doc, image_path, width_inches=3.0):
    """
    Insert an image into the document, centered, with a fixed width.
    """
    try:
        pic = doc.add_picture(image_path, width=Inches(width_inches))
        paragraph = doc.paragraphs[-1]
        paragraph.alignment = 1  # 0=left, 1=center, 2=right, 3=justify
    except Exception as e:
        print(f"Could not insert image {image_path}: {e}")

# Clinic reference image toggles
img_right_plagio = tk.BooleanVar()
img_left_plagio = tk.BooleanVar()
img_temporal_rotation = tk.BooleanVar()
img_pscb_1 = tk.BooleanVar()
img_pscb_2 = tk.BooleanVar()
img_parietal_bone = tk.BooleanVar()
img_sagittal_suture = tk.BooleanVar()
img_normal_lip_frenulum = tk.BooleanVar()
img_normal_tongue_frenulum_1 = tk.BooleanVar()
img_normal_tongue_frenulum_2 = tk.BooleanVar()
img_eop = tk.BooleanVar()

tk.Label(images_frame, text="Clinic Reference Images", font=("Arial", 14, "bold")).pack(anchor="w", padx=20, pady=(10,5))

tk.Checkbutton(images_frame, text="Right Plagiocephaly", variable=img_right_plagio).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Left Plagiocephaly", variable=img_left_plagio).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Temporal Bone Rotation / Ear Flare", variable=img_temporal_rotation).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Supernumerary Cranial Bones (PSCB 1)", variable=img_pscb_1).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Supernumerary Cranial Bones (PSCB 2)", variable=img_pscb_2).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Parietal Bone (Cephalohematoma)", variable=img_parietal_bone).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Sagittal Suture (Cephalohematoma)", variable=img_sagittal_suture).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Normal Lip Frenulum", variable=img_normal_lip_frenulum).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Normal Tongue Frenulum (Example 1)", variable=img_normal_tongue_frenulum_1).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Normal Tongue Frenulum (Example 2)", variable=img_normal_tongue_frenulum_2).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="External Occipital Protuberance (EOP)", variable=img_eop).pack(anchor="w", padx=40)

tk.Label(images_frame, text="Patient Images", font=("Arial", 14, "bold")).pack(anchor="w", padx=20, pady=(20,5))

# Container for 2-column grid layout
images_grid = tk.Frame(images_frame)
images_grid.pack(fill="both", expand=True, padx=20, pady=10)


def create_drop_zone(parent, label_text, key):
    """
    Creates a drag-and-drop zone for patient images.
    Includes:
      - Label
      - Thumbnail preview
      - Clear button
    Stores the dropped file path in patient_images[key].
    """

    # Outer frame (do NOT pack/grid here)
    frame = tk.Frame(parent, bd=2, relief="groove", padx=10, pady=10)

    # Label
    label = tk.Label(frame, text=label_text, width=50, height=2)
    label.pack()

    # Thumbnail preview area
    preview_label = tk.Label(frame)
    preview_label.pack(pady=5)

    # Clear button
    def clear_image():
        patient_images[key] = None
        label.config(text=label_text)
        preview_label.config(image="", text="")

    clear_btn = tk.Button(frame, text="Clear Image", command=clear_image)
    clear_btn.pack(pady=5)

    # Drag-and-drop registration
    label.drop_target_register("DND_Files")

    def drop_handler(event):
        path = event.data.strip().strip("{}")
        patient_images[key] = path

        # Update label
        label.config(text=f"Loaded: {os.path.basename(path)}")

        # Load a small thumbnail — just enough to confirm the right image was dropped
        try:
            img = Image.open(path)
            img.thumbnail((100, 100))
            tk_img = ImageTk.PhotoImage(img)
            img.close()

            preview_label.image = tk_img
            preview_label.config(image=tk_img)

        except Exception as e:
            preview_label.config(text=f"Error loading image: {e}")

    label.dnd_bind("<<Drop>>", drop_handler)

# --- Description textbox ---
    desc_label = tk.Label(frame, text="Image Description:", font=("Arial", 9))
    desc_label.pack()

    desc_entry = tk.Entry(frame, width=40)
    desc_entry.pack(pady=3)

    def save_description(event=None):
        patient_image_descriptions[key] = desc_entry.get().strip()

    desc_entry.bind("<KeyRelease>", save_description)

    return frame



# Two-column layout
row = 0
col = 0

def add_zone(label, key):
    global row, col
    frame = create_drop_zone(images_grid, label, key)
    frame.grid(row=row, column=col, padx=20, pady=10, sticky="nsew")

    col += 1
    if col > 1:  # 2 columns
        col = 0
        row += 1

# Now add zones using the helper
add_zone("Drop Anterior View Photo Here", "anterior_view")
add_zone("Drop Left Profile Photo Here", "left_profile")
add_zone("Drop Right Profile Photo Here", "right_profile")
add_zone("Drop Posterior View Photo Here", "posterior_view")
add_zone("Drop Superior View Photo Here", "superior_view")

add_zone("Drop Plagiocephaly Photo Here", "plagio")
add_zone("Drop Cephalohematoma Photo Here", "cephalohematoma")
add_zone("Drop Dolichocephaly Photo Here", "dolichocephaly")
add_zone("Drop Brachycephaly Photo Here", "brachycephaly")
add_zone("Drop Torticollis Photo Here", "torticollis")
add_zone("Drop TOTs Photo Here", "tots")
add_zone("Drop Metopic Ridge Photo Here", "metopic_ridge")
add_zone("Drop Occipital Extension Photo Here", "occipital_extension")
add_zone("Drop Frontal Alignment Photo Here", "frontal_alignment")
add_zone("Drop Temporal Alignment Photo Here", "temporal_alignment")
add_zone("Drop Retrognathia Photo Here", "retrognathia")
add_zone("Drop Prognathia Photo Here", "prognathia")
add_zone("Drop Lip Frenulum Photo Here", "lip_frenulum")
add_zone("Drop Tongue Frenulum Photo Here", "tongue_frenulum")



##=====================
# Practitioner Notes Tab
##=====================

tk.Label(practitioner_frame, text="Additional Notes:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20,0))

Additional_Notes = tk.Text(practitioner_frame, width=80, height=3, wrap="word")
Additional_Notes.pack(anchor="w", padx=20, pady=5)
Additional_Notes.config(state="normal")

# Practitioner Prefix
tk.Label(practitioner_frame, text="Practitioner Prefix:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
practitioner_prefix = tk.StringVar()
tk.Entry(practitioner_frame, textvariable=practitioner_prefix, width=40).pack(anchor="w", padx=20)

# Practitioner First Name
tk.Label(practitioner_frame, text="Practitioner First Name:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
practitioner_first_name = tk.StringVar()
tk.Entry(practitioner_frame, textvariable=practitioner_first_name, width=40).pack(anchor="w", padx=20)

# Practitioner Last Name
tk.Label(practitioner_frame, text="Practitioner Last Name:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
practitioner_last_name = tk.StringVar()
tk.Entry(practitioner_frame, textvariable=practitioner_last_name, width=40).pack(anchor="w", padx=20)

# Practitioner Suffix
tk.Label(practitioner_frame, text="Practitioner Suffix:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
practitioner_suffix = tk.StringVar()
tk.Entry(practitioner_frame, textvariable=practitioner_suffix, width=40).pack(anchor="w", padx=20)


def _on_mousewheel_p(event):
    canvas_p.yview_scroll(int(-1 * (event.delta / 120)), "units")

def bind_mousewheel_to_patient():
    root.unbind_all("<MouseWheel>")
    canvas_p.bind("<MouseWheel>", _on_mousewheel_p)
    bind_mousewheel_recursive(form_frame, _on_mousewheel_p)

def _bind_active_tab_scroll(*_):
    canvas = tab_canvas_map.get(notebook.select())
    if canvas:
        root.unbind_all("<MouseWheel>")
        root.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

def bind_mousewheel_to_practitioner():
    _bind_active_tab_scroll()

notebook.bind("<<NotebookTabChanged>>", _bind_active_tab_scroll)


# Patient First Name
tk.Label(form_frame, text="Patient First Name:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
patient_first_name = tk.StringVar()
tk.Entry(form_frame, textvariable=patient_first_name, width=40).pack(anchor="w", padx=20)

# Patient Last Name
tk.Label(form_frame, text="Patient Last Name:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
patient_last_name = tk.StringVar()
tk.Entry(form_frame, textvariable=patient_last_name, width=40).pack(anchor="w", padx=20)

# Date of Birth
tk.Label(form_frame, text="Date of Birth (mm/dd/yyyy):", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
patient_dob = tk.StringVar()
tk.Entry(form_frame, textvariable=patient_dob, width=40).pack(anchor="w", padx=20)

# Patient Gender (Dropdown)
tk.Label(form_frame, text="Gender:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
patient_gender = tk.StringVar()
gender_dropdown = ttk.Combobox(form_frame, textvariable=patient_gender, values=["Choose An Option", "Male", "Female"], width=37)
gender_dropdown.pack(anchor="w", padx=20)
gender_dropdown.current(0)

# Pronoun Variables
pronoun_subj = tk.StringVar()       # he / she
pronoun_subj_cap = tk.StringVar()   # He / She
pronoun_obj = tk.StringVar()        # him / her
pronoun_obj_cap = tk.StringVar()    # Him / Her
pronoun_poss = tk.StringVar()       # his / her
pronoun_poss_cap = tk.StringVar()   # His / Her

def update_pronouns(event=None):
    gender = patient_gender.get()

    if gender == "Male":
        pronoun_subj.set("he")
        pronoun_subj_cap.set("He")
        pronoun_obj.set("him")
        pronoun_obj_cap.set("Him")
        pronoun_poss.set("his")
        pronoun_poss_cap.set("His")

    elif gender == "Female":
        pronoun_subj.set("she")
        pronoun_subj_cap.set("She")
        pronoun_obj.set("her")
        pronoun_obj_cap.set("Her")
        pronoun_poss.set("her")
        pronoun_poss_cap.set("Her")

    else:
        # FIXED: no default male pronouns
        pronoun_subj.set("")
        pronoun_subj_cap.set("")
        pronoun_obj.set("")
        pronoun_obj_cap.set("")
        pronoun_poss.set("")
        pronoun_poss_cap.set("")

gender_dropdown.bind("<<ComboboxSelected>>", update_pronouns)

# IVF yes/no
tk.Label(form_frame, text="Conceived by IVF?", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
ivf_conception = tk.StringVar()
ivf_conception_dropdown = ttk.Combobox(form_frame, textvariable=ivf_conception, values=["Choose An Option", "Yes", "No"], width=37)
ivf_conception_dropdown.pack(anchor="w", padx=20)
ivf_conception_dropdown.current(0)

# Unmedicated?
tk.Label(form_frame, text="Was your birth unmedicated?", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
unmedicated = tk.StringVar()
unmedicated_dropdown = ttk.Combobox(form_frame, textvariable=unmedicated, values=["Choose An Option", "Yes (skip the next 2 questions)", "No"], width=37)
unmedicated_dropdown.pack(anchor="w", padx=20)
unmedicated_dropdown.current(0)

# Birth induction
tk.Label(form_frame, text="Was the birth medically induced?", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
induced = tk.StringVar()
induced_dropdown = ttk.Combobox(form_frame, textvariable=induced, values=["Choose An Option", "Yes", "No"], width=37)
induced_dropdown.pack(anchor="w", padx=20)
induced_dropdown.current(0)

# Epidural yes/no
tk.Label(form_frame, text="Did you receive an Epidural?", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
epidural = tk.StringVar()
epidural_dropdown = ttk.Combobox(form_frame, textvariable=epidural, values=["Choose An Option", "Yes", "No"], width=37)
epidural_dropdown.pack(anchor="w", padx=20)
epidural_dropdown.current(0)

# Delivery Mode
tk.Label(form_frame, text="Delivery Mode:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
birth_mode = tk.StringVar()
birth_mode_dropdown = ttk.Combobox(form_frame, textvariable=birth_mode, values=["Choose An Option", "Vaginal-unassisted", "Vaginal-instrument assisted", "Vaginal-manually assisted", "Caesarian Section"], width=37)
birth_mode_dropdown.pack(anchor="w", padx=20)
birth_mode_dropdown.current(0)

# Weeks at Birth
tk.Label(form_frame, text="Weeks at Birth (e.g., 39 weeks 2 days):", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
weeks_at_birth = tk.StringVar()
tk.Entry(form_frame, textvariable=weeks_at_birth, width=40).pack(anchor="w", padx=20)

# Internal Fetal Monitor
tk.Label(form_frame, text="Internal Fetal Monitor used?", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
ifm = tk.StringVar()
ifm_dropdown = ttk.Combobox(form_frame, textvariable=ifm, values=["Choose An Option", "Yes", "No"], width=37)
ifm_dropdown.pack(anchor="w", padx=20)
ifm_dropdown.current(0)

# Labor Length
tk.Label(form_frame, text="Length of Labor (hours):", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20,0))
labor_length = tk.StringVar()
tk.Entry(form_frame, textvariable=labor_length, width=40).pack(anchor="w", padx=20)

# Pushing Length
tk.Label(form_frame, text="Length of Pushing (hours):", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20,0))
pushing = tk.StringVar()
tk.Entry(form_frame, textvariable=pushing, width=40).pack(anchor="w", padx=20)

# complications

tk.Label(form_frame, text="Were there complications?", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
complications = tk.StringVar()
complications_dropdown = ttk.Combobox(form_frame, textvariable=complications, values=["Choose An Option", "Yes", "No"], width=37)
complications_dropdown.pack(anchor="w", padx=20)
complications_dropdown.current(0)

tk.Label(form_frame, text="If yes, describe:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(10, 0))
complications_text = tk.StringVar()
tk.Entry(form_frame, textvariable=complications_text, width=40).pack(anchor="w", padx=20)

# Birth Weight
tk.Label(form_frame, text="Birth Weight (lbs/oz):", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
birth_weight = tk.StringVar()
tk.Entry(form_frame, textvariable=birth_weight, width=40).pack(anchor="w", padx=20)

# Initial Breastfeeding
tk.Label(form_frame, text="Initial Breastfeeding Duration (hours/minutes):", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
breastfeeding_duration_initial = tk.StringVar()
tk.Entry(form_frame, textvariable=breastfeeding_duration_initial, width=40).pack(anchor="w", padx=20)

# main complaint
tk.Label(form_frame, text="Patient's main complaint", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
main_complaint = tk.StringVar()
tk.Entry(form_frame, textvariable=main_complaint, width=40).pack(anchor="w", padx=20)

# current feeding method
tk.Label(form_frame, text="Current Feeding Method(s):", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))

feeding_breast = tk.BooleanVar()
feeding_breast_bottle = tk.BooleanVar()
feeding_solid = tk.BooleanVar()

tk.Checkbutton(form_frame, text="Breastfeeding", variable=feeding_breast).pack(anchor="w", padx=40)
tk.Checkbutton(form_frame, text="Pumped Breastmilk via bottle", variable=feeding_breast_bottle).pack(anchor="w", padx=40)
tk.Checkbutton(form_frame, text="Solid or pureed foods", variable=feeding_solid).pack(anchor="w", padx=40)

# Current Breastfeeding Duration
tk.Label(form_frame, text="Breastfeeding Habits (e.g.'for the first 3 months' or 'until present' or 'none'):", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
breastfeeding_duration_current = tk.StringVar()
tk.Entry(form_frame, textvariable=breastfeeding_duration_current, width=40).pack(anchor="w", padx=20)

# breastfeeding complaints
tk.Label(form_frame, text="Please list any breastfeeding complaints (e.g. nipple soreness) or enter 'None'", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
breast_complaint = tk.StringVar()
tk.Entry(form_frame, textvariable=breast_complaint, width=40).pack(anchor="w", padx=20)

# Breastfeeding Digestive Issues
tk.Label(form_frame, text="Noted Digestive Issues with Breastfeeding (e.g., colic, trapped gas, reflux, etc.):", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
breast_digestive_issues = tk.StringVar()
tk.Entry(form_frame, textvariable=breast_digestive_issues, width=40).pack(anchor="w", padx=20)

# Formula Use
tk.Label(form_frame, text="Formula feeding:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
formula_use = tk.StringVar()
formula_use_dropdown = ttk.Combobox(form_frame, textvariable=formula_use, values=["None", "Primary feeding method", "Supplemental only"], width=37)
formula_use_dropdown.pack(anchor="w", padx=20)
formula_use_dropdown.current(0)

# Formula type
tk.Label(form_frame, text="Formula Brand:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
formula_type = tk.StringVar()
tk.Entry(form_frame, textvariable=formula_type, width=40).pack(anchor="w", padx=20)

# Digestive Issues
tk.Label(form_frame, text="Noted Digestive Issues with Formula (gas, reflux, constipation, etc.):", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
digestive_issues = tk.StringVar()
tk.Entry(form_frame, textvariable=digestive_issues, width=40).pack(anchor="w", padx=20)


# Solids introduced?
tk.Label(form_frame, text="Has the patient started solids?", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
solids = tk.StringVar()
solids_dropdown = ttk.Combobox(form_frame, textvariable=solids, values=["Choose An Option", "Yes", "No (Skip the next question)"], width=37)
solids_dropdown.pack(anchor="w", padx=20)
solids_dropdown.current(0)

# Age Solids introduced
tk.Label(form_frame, text="Age Solids introduced:", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
solids_age = tk.StringVar()
tk.Entry(form_frame, textvariable=solids_age, width=40).pack(anchor="w", padx=20)

# Previous Care
tk.Label(form_frame, text="Has the patient received previous care for this condition?", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
previous_care = tk.StringVar()
previous_care_dropdown = ttk.Combobox(form_frame, textvariable=previous_care, values=["Choose An Option", "Yes", "No (Skip the next question)"], width=37)
previous_care_dropdown.pack(anchor="w", padx=20, pady=(10, 0))
previous_care_dropdown.current(0)

# Other Care Details
tk.Label(form_frame, text="What other care has the patient received for this condition?", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
other_care_details = tk.StringVar()
tk.Entry(form_frame, textvariable=other_care_details, width=40).pack(anchor="w", padx=20, pady=(10, 0))

# Submit button
def submit_data():
    print("Patient First Name:", patient_first_name.get())
    print("Patient Last Name:", patient_last_name.get())
    print("Date of Birth:", patient_dob.get())
    print("Gender:", patient_gender.get())
    print("Pronoun (subj):", pronoun_subj.get())
    print("Pronoun (subj-cap):", pronoun_subj_cap.get())
    print("Pronoun (obj):", pronoun_obj.get())
    print("Pronoun (obj-cap):", pronoun_obj_cap.get())
    print("Pronoun (poss):", pronoun_poss.get())
    print("Pronoun (poss-cap):", pronoun_poss_cap.get())
    print("IVF:", ivf_conception.get())
    print("Unmedicated?:", unmedicated.get())
    print("Birth induction:", induced.get())
    print("Epidural:", epidural.get())
    print("Birth mode:", birth_mode.get())
    print("Weeks at Birth:", weeks_at_birth.get())
    print("IFM:", ifm.get())
    print("Labor Length:", labor_length.get())
    print("Pushing Length:", pushing.get())
    print("Complications:", complications.get())
    print("Complication Details:", complications_text.get())
    print("Birth Weight:", birth_weight.get())
    print("Initial Breastfeeding Duration:", breastfeeding_duration_initial.get())
    print("Current Breastfeeding Duration:", breastfeeding_duration_current.get())
    print("Feeding Method - Breastfeeding:", feeding_breast.get())
    print("Feeding Method - Breastmilk Bottle:", feeding_breast_bottle.get())
    print("Feeding Method - Solids:", feeding_solid.get())
    print("Formula Use:", formula_use.get())
    print("Formula Type:", formula_type.get())
    print("Digestive Issues:", digestive_issues.get())
    print("Solids introduced?", solids.get())
    print("Age Solids introduced:", solids_age.get())
    print("Breastfeeding Complaints:", breast_complaint.get())
    print("Patient's Main Complaint:", main_complaint.get())
    print("Noted Digestive Issues with Breastfeeding:", breast_digestive_issues.get())
    print("Previous Care:", previous_care.get())
    print("Other Care Details:", other_care_details.get())
    print("Additional Notes:", Additional_Notes.get("1.0", "end").strip())
    print("Practitioner Prefix:", practitioner_prefix.get())
    print("Practitioner First Name:", practitioner_first_name.get())
    print("Practitioner Last Name:", practitioner_last_name.get())
    print("Practitioner Suffix:", practitioner_suffix.get())
    print("=========================================")
    print("PatientData submission complete.")
    print("=========================================")
    print("")
    print("Practitioner Findings:")
    print("Cranial - Cephalohematoma:", cephalohematoma.get())
    print("Cephalohematoma Measurement:", cephalohematoma_measurement.get())
    print("Cephalohematoma Location:", cephalohematoma_location.get())
    print("Cephalohematoma Cross Suture Lines:", cephalohematoma_cross_suture.get())
    print("Cephalohematoma:", cephalohematoma_calcium.get())
    print("Cranial - Brachycephaly:", brachycephaly.get())
    print("Cranial - Bracycephaly Severity:", brachycephaly_severity.get())
    print("Cranial - Bracycephaly Measurement:", brachycephaly_measurement.get())
    print("Cranial - Plagiocephaly:", plagiocephaly.get())
    print("Cranial - Plagiocephaly Severity:", plagiocephaly_severity.get())
    print("Cranial - Plagiocephaly Measurement:", plagiocephaly_measurement.get())
    print("Laterality:", plagiocephaly_laterality.get())
    print("Cranial - Dolichocephaly:", dolichocephaly.get())
    print("Cranial - Dolichocephaly Severity:", dolichocephaly_severity.get())
    print("Cranial - Dolichocephaly Measurement:", dolichocephaly_measurement.get())
    print("Facial Asymmetry:", facial_asymmetry.get())
    print("Occipital Extension:", occipital_extension.get())
    print("Torticollis:", torticollis.get())
    print("Torticollis Laterality:", torticollis_laterality.get())
    print("Tethered Oral Tissues (TOTs):", tethered_oral_tissues.get())
    print("Internal Maxilla Rotation:", internal_maxilla_rotation.get())
    print("Retrognathia:", retrognathia.get())
    print("Prognathia:", prognathia.get())
    print("Thrush:", thrush.get())
    print("Cradle Cap:", cradle_cap.get())
    print("Blocked Tear Duct:", blocked_tear_duct.get())
    print("Blocked Tear Duct Laterality:", blocked_tear_duct_laterality.get())
    print("Acetabular Click:", acetabular_click.get())
    print("Acetabular Click Laterality:", acetabular_click_laterality.get())
    print("Femur Rotation:", femur_rotation.get())
    print("Femur Rotation Laterality:", femur_rotation_laterality.get())
    print("Femur Rotation Restriction:", femur_rotation_restriction.get())
    print("=========================================")
    print("End of Practitioner Findings")
    print("=========================================")
    print("")
    print("Recommendations:")
    print("Thrush Recommendations:", thrush_rec_box.get("1.0", "end").strip())
    print("Cradle Cap Recommendations:", cradle_cap_rec_box.get("1.0", "end").strip())
    print("Blocked Tear Duct Recommendations:", blocked_tear_duct_rec_box.get("1.0", "end").strip())
    print("Cephalohematoma Recommendations:", cephalohematoma_rec_box.get("1.0", "end").strip())
    print("Jaw Exercise Recommendations:", jaw_ex_box.get("1.0", "end").strip())
    print("Homeopathy Recommendations:", homeo_rec_box.get("1.0", "end").strip())
    print("Homeopathy Consultation Recommendations:", homeo_cons_rec_box.get("1.0", "end").strip())
    print("Primitive Reflex Recommendations:", primitive_reflex_rec_box.get("1.0", "end").strip())
    print("Observed Primitive Reflexes:", primitive_reflexes_findings.get("1.0", "end").strip())
    print("TOTs Dentist Recommendations:", tots_docs_rec_box.get("1.0", "end").strip())
    print("==========================================")
    print("")
    print("Practitioner Notes:")
    print("Additional Notes:", Additional_Notes.get("1.0", "end").strip())
    print("Practitioner Name:", practitioner_prefix.get(), practitioner_first_name.get(), practitioner_last_name.get(), practitioner_suffix.get())
    print("==========================================")
    print("End of Submission")
    print("==========================================")


# --- Bottom Submit Area for Patient Page ---
patient_bottom = tk.Frame(form_frame)
patient_bottom.pack(fill="x", pady=30)

tk.Button(patient_bottom, text="Submit Patient Info", command=submit_data).pack()
tk.Button(patient_bottom, text="Practitioner Mode", command=show_practitioner_frame).pack(pady=10)
tk.Button(patient_bottom, text="Save Progress", command=lambda: save_progress(), bg="#2196F3", fg="white").pack(pady=2)
tk.Button(patient_bottom, text="Load Saved Progress", command=lambda: load_progress(), bg="#2196F3", fg="white").pack(pady=2)  

# --- Bottom bar for Practitioner Page (fixed, always visible) ---
practitioner_bottom = tk.Frame(frame_practitioner)
practitioner_bottom.pack(fill="x", pady=5)

tk.Button(practitioner_bottom, text="Submit All Data", command=submit_data).pack()
tk.Button(practitioner_bottom, text="Back to Patient Mode", command=show_patient_frame).pack(pady=5)
tk.Button(practitioner_bottom, text="Save Progress", command=lambda: save_progress(), bg="#2196F3", fg="white").pack(pady=2)
tk.Button(practitioner_bottom, text="Load Saved Progress", command=lambda: load_progress(), bg="#2196F3", fg="white").pack(pady=2)




# &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
# ---------- Save / Load Progress --------------------------
# &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

# All StringVar names in the form
_STRING_VAR_NAMES = [
    "date_of_service",
    "patient_first_name", "patient_last_name", "patient_dob", "patient_gender",
    "ivf_conception", "unmedicated", "induced", "epidural", "birth_mode", "weeks_at_birth",
    "ifm", "labor_length", "pushing", "complications", "complications_text",
    "birth_weight", "breastfeeding_duration_initial", "breastfeeding_duration_current",
    "main_complaint", "breast_complaint", "breast_digestive_issues",
    "formula_use", "formula_type", "digestive_issues", "solids", "solids_age",
    "previous_care", "other_care_details", "treatment_plan",
    "practitioner_prefix", "practitioner_first_name", "practitioner_last_name", "practitioner_suffix",
    "cephalohematoma_measurement", "cephalohematoma_location",
    "cephalohematoma_cross_suture", "cephalohematoma_calcium", "cephalohematoma_calcium_location",
    "brachycephaly_severity", "brachycephaly_measurement",
    "plagiocephaly_severity", "plagiocephaly_measurement", "plagiocephaly_laterality",
    "dolichocephaly_severity", "dolichocephaly_measurement",
    "torticollis_laterality", "blocked_tear_duct_laterality",
    "acetabular_click_laterality", "femur_rotation_laterality", "femur_rotation_restriction",
    "temporal_ear_flare_laterality", "temporal_alignment_outcome",
]

# All BooleanVar names in the form
_BOOL_VAR_NAMES = [
    "cephalohematoma", "brachycephaly", "plagiocephaly", "dolichocephaly",
    "facial_asymmetry", "occipital_extension", "sutural_restriction", "metopic_ridge", "supernumerary_bones",
    "frontal_alignment", "temporal_alignment",
    "torticollis", "tethered_oral_tissues", "internal_maxilla_rotation",
    "retrognathia", "prognathia",
    "thrush", "cradle_cap", "blocked_tear_duct", "acetabular_click",
    "femur_rotation", "inappropriate_reflexes", "skeletal_alignment",
    "homeo_rec", "homeo_cons_rec", "primitive_reflex_rec",
    "feeding_breast", "feeding_breast_bottle", "feeding_solid",
    "img_right_plagio", "img_left_plagio", "img_temporal_rotation",
    "img_pscb_1", "img_pscb_2", "img_parietal_bone", "img_sagittal_suture",
    "img_normal_lip_frenulum",
    "img_normal_tongue_frenulum_1", "img_normal_tongue_frenulum_2", "img_eop",
]

def _get_text_widget_map():
    return {
        "Additional_Notes": Additional_Notes,
        "thrush_rec_box": thrush_rec_box,
        "cradle_cap_rec_box": cradle_cap_rec_box,
        "cephalohematoma_rec_box": cephalohematoma_rec_box,
        "blocked_tear_duct_rec_box": blocked_tear_duct_rec_box,
        "jaw_ex_box": jaw_ex_box,
        "homeo_rec_box": homeo_rec_box,
        "homeo_cons_rec_box": homeo_cons_rec_box,
        "primitive_reflex_rec_box": primitive_reflex_rec_box,
        "tots_docs_rec_box": tots_docs_rec_box,
        "primitive_reflexes_findings": primitive_reflexes_findings,
    }

def _run_all_toggles():
    """Re-run every toggle function so dependent fields reflect current checkbox state."""
    toggle_cephalohematoma_fields()
    toggle_brachycephaly_fields()
    toggle_plagiocephaly_fields()
    update_plagio_image()
    toggle_supernumerary_bones()
    toggle_dolichocephaly_fields()
    toggle_temporal_alignment_fields()
    toggle_torticollis_fields()
    toggle_tethered_oral_tissues_fields()
    update_jaw_exercises()
    toggle_thrush_fields()
    toggle_cradle_cap_fields()
    toggle_blocked_tear_duct_fields()
    toggle_acetabular_click_fields()
    toggle_femur_rotation_fields()
    toggle_homeopathy()
    toggle_homeo_consult()
    toggle_primitive_reflex()
    update_pronouns()

def save_progress():
    filepath = filedialog.asksaveasfilename(
        defaultextension=".json",
        filetypes=[("Report Save File", "*.json"), ("All Files", "*.*")],
        title="Save Progress",
        initialfile="report_in_progress",
    )
    if not filepath:
        return

    data = {}

    # StringVars
    data["string_vars"] = {}
    for name in _STRING_VAR_NAMES:
        var = globals().get(name)
        if isinstance(var, tk.StringVar):
            data["string_vars"][name] = var.get()

    # BooleanVars
    data["bool_vars"] = {}
    for name in _BOOL_VAR_NAMES:
        var = globals().get(name)
        if isinstance(var, tk.BooleanVar):
            data["bool_vars"][name] = var.get()

    # Text widgets (temporarily enable to read, then restore state)
    data["text_widgets"] = {}
    for name, widget in _get_text_widget_map().items():
        try:
            original_state = widget.cget("state")
            widget.config(state="normal")
            data["text_widgets"][name] = widget.get("1.0", "end-1c")
            widget.config(state=original_state)
        except Exception:
            data["text_widgets"][name] = ""

    # Patient image paths and descriptions
    data["patient_images"] = {k: v for k, v in patient_images.items()}
    data["patient_image_descriptions"] = dict(patient_image_descriptions)

    try:
        with open(filepath, "w") as f:
            json.dump(data, f, indent=2)
        messagebox.showinfo("Saved", f"Progress saved to:\n{filepath}")
    except Exception as e:
        messagebox.showerror("Save Error", str(e))


def load_progress():
    filepath = filedialog.askopenfilename(
        filetypes=[("Report Save File", "*.json"), ("All Files", "*.*")],
        title="Load Saved Progress",
    )
    if not filepath:
        return

    try:
        with open(filepath, "r") as f:
            data = json.load(f)
    except Exception as e:
        messagebox.showerror("Load Error", f"Could not read file:\n{e}")
        return

    # 1. Restore StringVars
    for name, value in data.get("string_vars", {}).items():
        var = globals().get(name)
        if isinstance(var, tk.StringVar):
            var.set(value)

    # 2. Restore BooleanVars
    for name, value in data.get("bool_vars", {}).items():
        var = globals().get(name)
        if isinstance(var, tk.BooleanVar):
            var.set(value)

    # 3. Run all toggles — this sets the correct enabled/disabled state on every
    #    dependent field, and auto-fills recommendation boxes with their defaults.
    _run_all_toggles()

    # 4. Overwrite recommendation boxes with the actual saved text.
    #    We temporarily enable each widget to write to it, then restore its state.
    for name, content in data.get("text_widgets", {}).items():
        widget = _get_text_widget_map().get(name)
        if widget is None:
            continue
        current_state = widget.cget("state")
        widget.config(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", content)
        widget.config(state=current_state)

    # 5. Restore patient image paths and descriptions
    patient_images.update(data.get("patient_images", {}))
    patient_image_descriptions.update(data.get("patient_image_descriptions", {}))

    messagebox.showinfo(
        "Loaded",
        "Progress loaded successfully.\n\n"
        "Note: image thumbnails will not re-display, but any saved image paths "
        "will still be included when you generate the report."
    )


# &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
# ---------- Begin: Generate Report Functionality ----------
# &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

# Helper: safe get text from Text widget (handles disabled state)
def get_text_widget_content(widget):
    try:
        return widget.get("1.0", "end").strip()
    except Exception:
        return ""

def replace_text_in_paragraph(paragraph, placeholder, new_text):
    """
    Robust placeholder replacement.

    This handles cases where Word splits a placeholder across multiple runs.
    It preserves the paragraph's general style, but the replaced paragraph text
    will inherit the formatting of the first run.
    """
    full_text = "".join(run.text for run in paragraph.runs)

    if placeholder not in full_text:
        return False

    new_full_text = full_text.replace(placeholder, str(new_text))

    if paragraph.runs:
        paragraph.runs[0].text = new_full_text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(new_full_text)

    return True


def replace_placeholder(doc, placeholder, new_text):
    """
    Replace a placeholder in:
    - body paragraphs
    - table cells
    - headers
    - footers
    """

    # Main body paragraphs
    for p in doc.paragraphs:
        if placeholder in p.text:
            replace_text_in_paragraph(p, placeholder, new_text)

    # Main body tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if placeholder in p.text:
                        replace_text_in_paragraph(p, placeholder, new_text)

    # Headers and footers
    for section in doc.sections:
        header = section.header
        footer = section.footer

        for p in header.paragraphs:
            if placeholder in p.text:
                replace_text_in_paragraph(p, placeholder, new_text)

        for table in header.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if placeholder in p.text:
                            replace_text_in_paragraph(p, placeholder, new_text)

        for p in footer.paragraphs:
            if placeholder in p.text:
                replace_text_in_paragraph(p, placeholder, new_text)

        for table in footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if placeholder in p.text:
                            replace_text_in_paragraph(p, placeholder, new_text)

# Helper: replace placeholder with multiple paragraphs (for obj_fin entries)
def replace_placeholder_with_paragraphs(doc, placeholder, paragraphs_list):
    # Search paragraphs
    for i, p in enumerate(list(doc.paragraphs)):
        if placeholder in p.text:
            parent = p._p.getparent()
            # Replace the paragraph text with the first paragraph
            first_text = paragraphs_list[0] if paragraphs_list else ""
            replace_text_in_paragraph(p, placeholder, first_text)
            # Insert remaining paragraphs after this one
            insert_index = i + 1
            for extra in paragraphs_list[1:]:
                new_p = doc.add_paragraph(extra)
                # Move new_p to correct position by reordering XML
                parent.insert(insert_index, new_p._p)
                insert_index += 1
            return True
    # Search in tables if not found in top-level paragraphs
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for i, p in enumerate(list(cell.paragraphs)):
                    if placeholder in p.text:
                        replace_text_in_paragraph(p, placeholder, paragraphs_list[0] if paragraphs_list else "")
                        insert_index = i + 1
                        for extra in paragraphs_list[1:]:
                            new_p = cell.add_paragraph(extra)
                            # move to correct position
                            cell._tc.insert(insert_index, new_p._p)
                            insert_index += 1
                        return True
    return False

# Build pronouns from patient_gender
def build_pronouns(gender_value):
    g = (gender_value or "").strip().lower()
    if g.startswith("m"):
        return {
            "pronoun_subj": "he",
            "pronoun_subj_cap": "He",
            "pronoun_obj": "him",
            "pronoun_obj_cap": "Him",
            "pronoun_poss": "his",
            "pronoun_poss_cap": "His",
        }
    elif g.startswith("f"):
        return {
            "pronoun_subj": "she",
            "pronoun_subj_cap": "She",
            "pronoun_obj": "her",
            "pronoun_obj_cap": "Her",
            "pronoun_poss": "her",
            "pronoun_poss_cap": "Her",
        }
    else:
        return {
            "pronoun_subj": "they",
            "pronoun_subj_cap": "They",
            "pronoun_obj": "them",
            "pronoun_obj_cap": "Them",
            "pronoun_poss": "their",
            "pronoun_poss_cap": "Their",
        }

# Objective finding templates
finding_templates = {
"cephalohematoma": lambda ctx: (
    f"Cephalohematoma – {ctx['patient_name']}'s cephalohematoma measured "
    f"{ctx.get('cephalohematoma_measurement','___')} and was located {ctx.get('cephalohematoma_location','___')}. "
    f"{'It crossed the sutural lines' if is_yes(ctx.get('cephalohematoma_cross_suture','')) else 'It did not cross the sutural lines'} "
    f"and has calcium deposits {ctx.get('cephalohematoma_calcium_location','___')}. "
    f"A cephalohematoma is often due to birth trauma like being stuck in transition too long or the baby's head moving in and out repeatedly during transition. "
    f"These swellings often resolve within days after birth, but when they take longer to resolve, there is an issue with efficiently routing the fluid through the cranial ventricular system. "
    f"Cranial treatments combined with homeopathy often give a wonderful result. "
    f"When fluid stagnates in the sagittal sinus along the motor homunculus, seizures can ensue — this can be resolved with cranial technique. "
    f"Occasionally, during resolution of the cephalohematoma, there will be bits of blood in the urine as the fluid is flushed out. "
    f"The homeopathic remedies help resolve calcium depositions."
    ),
"post_cephalohematoma_calcium_deposition": lambda ctx: (
    f"Post Cephalohematoma Calcium Deposition – {ctx['patient_name']}'s cephalohematoma had resolved at the time of "
    f"{ctx['pronoun_poss']} exam, but calcium deposits are palpable along the previous CH edges. "
    f"A cephalohematoma (CH) is often due to birth trauma like being stuck in transition too long or the baby's head moving in and out repeatedly during transition. "
    f"These swellings often resolve within days after birth, but when they take longer to resolve, there is an issue with efficient routing of the Cerebrospinal Fluid (CSF) through the cranial ventricular system. "
    f"Cranial treatments combined with homeopathy often give a wonderful result. "
    f"When CSF stagnates in the sagittal sinus along the motor homunculus, seizures can ensue. "
    f"It is important to have cranial treatment after the EXTERNAL resolution of the CH to make sure the CSF is moving efficiently INTERNALLY throughout the brain sinuses and ventricles. "
    f"The homeopathic remedies help to safely resolve calcium depositions."
    ),
"brachycephaly": lambda ctx: (
    f"Brachycephaly – {ctx['patient_name']} presents with a {ctx.get('brachycephaly_severity','___')} "
    f"non-synostotic brachycephaly. Non-synostotic means there is no sutural fusion or aplasia — in other words, the patient is not a candidate for cranial surgery. "
    f"Brachycephaly means the back of the head is flat and the front-to-back distance is too short as compared to the cranial width. "
    f"This is reflected in a Cephalic Index (CI) measurement of {ctx.get('brachycephaly_measurement','___')}. "
    f"The CI measures the proportionality of the length and width of the head — the preferred normal CI is closer to 80. "
    f"The back of the head houses the cerebellum, which coordinates primitive reflexes, muscle tone, balance and posture, as well as attention, language, and emotional regulation. "
    f"A flat occiput also increases the chance of scoliosis due to a compensatory reverse cervical curve."
    ),
"plagiocephaly": lambda ctx: (
    f"Plagiocephaly – {ctx['patient_name']} measures for {ctx.get('plagiocephaly_severity','___')} "
    f"non-synostotic plagiocephaly, with a Cranial Vault Asymmetry (CVA) measured at {ctx.get('plagiocephaly_measurement','___')} (0–1 is the desired measurement). "
    f"Non-synostotic means there is no sutural fusion or aplasia — the patient is not a candidate for cranial surgery. "
    f"Plagiocephaly means that one side of the head is not symmetrical with the other. "
    f"The {ctx.get('plagiocephaly_laterality','___')} side of {ctx['pronoun_poss']} cranial vault and/or face is anterior, or forward, as compared to the other side."
    + (
        f" {ctx['patient_first_name']} also has a restriction of the vertebrae in {ctx['pronoun_poss']} upper neck (torticollis) causing a preference for which way "
        f"{ctx['pronoun_poss']} head is comfortable. This preference can contribute to {ctx['pronoun_poss']} plagiocephaly, and the plagiocephaly sustains the misalignment in "
        f"{ctx['pronoun_poss']} neck — it can be a vicious cycle."
        if globals().get('torticollis') and globals()['torticollis'].get() else ""
    )
    ),
"dolichocephaly": lambda ctx: (
    f"Dolichocephaly – {ctx['patient_name']} presents with a {ctx.get('dolichocephaly_severity','___')} "
    f"non-synostotic dolichocephalic presentation. Non-synostotic means this is malleable — not a result of sutural fusion or aplasia, and not a surgical candidate. "
    f"{ctx['pronoun_poss_cap']} Cephalic Index (CI) measured at {ctx.get('dolichocephaly_measurement','___')}; ideal range is closer to 80. "
    f"A narrow cranium causes (1) a high palate which can make nursing more challenging, and (2) teeth crowding and a proclivity toward orthodontics later in life. "
    f"A high palate is especially problematic in the presence of a tongue tie, making extraction of milk at the breast incomplete, leading to low milk supply and an unsatisfied baby."
    ),
"torticollis": lambda ctx: (
    f"Torticollis – A mild to moderate {ctx.get('torticollis_laterality','___')} lateral torticollis was noted. "
    f"This is a fancy term for a crick in the neck, causing the neck to prefer turning to one side. "
    f"This can be due to an internal dural torsion, vertebral fixation/positioning, and/or muscular imbalance, and can sometimes cause nursing on one breast to be more comfortable than the other. "
    f"Home exercises of inversion swinging a couple of times daily will assist treatment by straightening this dural torsion. "
    f"A good lactation consultant will be able to help find comfortable positions for each breast. "
    f"Vagal nerve compression often accompanies torticollis and can cause digestive issues as well — the vagal nerve will recover as the torticollis is resolved."
    ),
"tethered_oral_tissues": lambda ctx: (
    f"Tethered Oral Tissues (TOTs) – {ctx['patient_first_name']}'s oral ties need to be evaluated for possible "
    f"laser frenectomy by a pediatric dentist extensively trained in TOTs. "
    f"Oral ties can cause a variety of nursing issues and breast pain. "
    f"TOTs can also cause incomplete extraction of milk, low weight gain, painful gas, diminishing milk supply, mastitis, and more. "
    f"Speech, dentition, neck/throat tension, posture, gag reflex, and airway function can all be negatively affected by not releasing ties during infancy. "
    f"TOTs can cause delayed response to cranial treatments. Infancy is by far the best time to have this done."
    ),
"internal_maxilla_rotation": lambda ctx: (
    f"Internal Maxilla Rotation – A high, narrow palate with a low intermaxillary suture was observed and is commonly the result of internally rotated maxillae and intermaxillary sutural restriction. "
    f"This is often seen as a result of a tongue tie. "
    f"The tongue's pressure on the top of the mouth helps shape the palate — if the tongue is not fully contacting the upper palate, the mouth shape will be off. "
    f"Incidentally, internal maxillae rotation often presents with dark circles under the eyes."
    ),
"sutural_restrictions": lambda ctx: (
    f"Sutural Restrictions – Cranial findings present with tight vault, oral, and facial sutures with ridging along several cranial vault sutures. "
    f"This is a common finding in C-sections, babies trapped in transition, or those with a transition that was too fast. "
    f"Sutural restriction limits cranial expansion. This restriction can involve just one area of the head causing asymmetry, limit the expansion of the sinus or oral structures, "
    f"or result in limited expansion of the whole cranial vault. "
    f"When considering it is brain growth that causes the cranium to expand, one can see the importance of having sutures that don't restrict cranial expansion. "
    f"Sometimes these asymmetries don't show up until the 12-week growth spurt. "
    f"Is this just aesthetics? No — the shape of the head can directly affect brain development, TMJ and airway function, and future orthodontic needs."
    ),
"metopic_ridging": lambda ctx: (
    f"Metopic Ridging – The metopic suture runs up and down the center of the forehead and is the only cranial suture that actually fuses. "
    f"This typically happens between 3–9 months of age. "
    f"Cranial therapy before this window can improve the ridging when the suture is jammed. "
    f"Cranial therapy after this time window can sometimes reduce the severity of the ridging over growth spurts, but takes time — sometimes months, sometimes years."
    ),
"potential_supernumerary_cranial_bones": lambda ctx: (
    f"Potential Supernumerary Cranial Bones – It is not uncommon to find normal variants in the cranial morphology, especially along the lambdoid suture. "
    f"These look like little bonus bones, and are of no functional or aesthetic consequence unless there is a flattening of the lambdoid suture with sutural restrictions. "
    f"We can usually get the sutures to release and the occiput and/or parietals to resume normal shape and positioning; however, it can be hard to get those little bonus bones to push back out. "
    f"This is of no functional consequence, and only occasionally do the little flat bone islands remain noticeable. "
    f"An x-ray can determine if these bones are present, but since it is of no real consequence, I have yet to opt for infant radiation solely for determining the presence or absence of such an inconsequential variant."
    ),
"skeletal_alignment": lambda ctx: (
    f"Skeletal Alignment – Misalignments or joint restrictions were discovered in the cervical, thoracic, lumbar, and sacral regions which can have associated neurological interference. "
    f"This can be due to external vertebral restrictions, muscle imbalance, or dural torsion. "
    f"These will be easily resolved at {ctx['pronoun_poss']} visits, along with the inversion swinging exercises."
    ),
"acetabular_click": lambda ctx: (
    f"Acetabular Click – An acetabular click was noticed on the {ctx.get('acetabular_click_laterality','___')} hip. "
    f"Acetabular x-rays are recommended."
    ),
"femur_rotation": lambda ctx: (
    f"Femur Rotation – Decreased {ctx.get('femur_rotation_restriction','___')} rotation of the "
    f"{ctx.get('femur_rotation_laterality','___')} femur was noted. "
    f"This may reflect muscular imbalance, pelvic torsion, or intrauterine positioning patterns. "
    f"Cranial and pelvic balancing often improve femoral rotation symmetry."
    ),
"thrush": lambda ctx: (
    f"Thrush – Thrush was visibly apparent on the tongue, which can cause painful nursing for the infant's mouth at the breast, bloating, reflux, and more. "
    f"Thrush is a fungal infection that can be local or systemic. "
    f"A high-quality probiotic (SFI Lacto Prime or SFI Baby Probiotic) is recommended twice daily. "
    f"Anovite Colostrum is recommended for baby. "
    f"If breastfeeding, SFI Ther-biotic Complete Probiotics and eliminating sugar/simple carbs may support maternal flora. "
    f"Nipples may be sprayed with Argentyn 23 colloidal silver before and after nursing, and Silverette Nipple Shields are recommended. "
    f"If symptoms do not resolve within two weeks, a consultation with Stephanie Kononovich (https://foundationhomeopathy.com) is recommended. "
    f"This should be resolved before proceeding with any TOTs frenectomies."
    ),
"cradle_cap": lambda ctx: (
    f"Cradle Cap – Cradle cap was noted. This is a fungal condition in the presence of a weakened immune system. "
    f"Please avoid scrubbing the flakes off, as this can cause the condition to spread. "
    f"Argentyn 23 silver gel or spray can be applied to the area, followed with massaging coconut or olive oil into the scalp. "
    f"If there is no improvement in 3–4 weeks, a consultation with Stephanie Kononovich (http://foundationhomeopathy.com) for a more specific homeopathic recommendation is advised."
    ),
"blocked_tear_duct": lambda ctx: (
    f"Blocked Tear Duct – A blocked {ctx.get('blocked_tear_duct_laterality','___')} tear duct was noted. "
    f"The homeopathic cell salt Silicea 6x is recommended 2–3x daily. "
    f"Argentyn 23 silver (spray or drops) in the eye 2x daily is also recommended. "
    f"Sutural release may help as well."
    ),
"facial_asymmetry": lambda ctx: (
    f"Facial Asymmetry – Facial asymmetry was noted. This may reflect underlying sutural restriction, "
    f"positional preference, or muscular imbalance. Cranial work typically improves symmetry over growth spurts."
    ),
"occipital_extension": lambda ctx: (
    f"Occipital Extension – This presents as a long, narrow head shape (similar to dolichocephaly) with a rounded, high subocciput. "
    f"It may reflect occipital bone positioning or dural tension patterns. Cranial work helps restore balanced occipital motion."
    ),
"frontal_alignment": lambda ctx: (
    f"Frontal Alignment – An externally rotated frontal bone was noted. This may present as a high brow, wider eye spacing, "
    f"anterior malar positioning, or chin deviation. Although this may seem only aesthetic, it can affect TMJ function, vision, and maxillary rotation."
    ),
"temporal_alignment": lambda ctx: (
    f"Temporal Alignment – A {ctx.get('temporal_ear_flare_laterality','___')} ear flare was noticed. "
    f"This is generally due to an external rotation of one or both temporal bones — the bone surrounding the ear. "
    f"Uncorrected, this can contribute to a crossbite, recessed jaw, more sensitive ears, motion sickness, ear infections, vertigo or balance issues, and auditory processing difficulties."
    ),
"retrognathia": lambda ctx: (
    f"Retrognathia – This is a fancy term for a recessed chin, which can later become an overbite. "
    f"This tendency is noticeable by the large crease under the lower lip and the flared temporal bones (above the ears). "
    f"This is often due to externally rotated temporal bones and can affect the position of the ear canal and the TM joint. "
    f"A recessed jaw can cause nursing pain and can be an orthodontic challenge later in life. Jaw exercises have been recommended."
    ),
"prognathia": lambda ctx: (
    f"Prognathia – This is a fancy term for an underbite. "
    f"This can be due to a variety of factors including, but not limited to, the internal rotation of the temporal bones causing the jaw to jut forward, mouth breathing, and prolonged thumb-sucking."
    ),
"inappropriate_reflexes": lambda ctx: (
    f"Inappropriately Exhibited Primitive Reflexes – Primitive reflexes were observed including: {ctx.get('primitive_reflex_findings','___')}. "
    f"My treatment results can be delayed due to inappropriate Primitive Reflexes, not to mention the challenges in daily life. "
    f"These reflexes can be easily organized and calmed with specific exercises. "
    f"Babies can then enjoy their environments without fear and stress. "
    f"Disorganized reflexes can cause asymmetrical tension in the body and disorganization in the brain. "
    f"Peak synaptic development is at 8 months. The cerebellum increases in size by 240% during the first year of life. "
    f"By 2 years the brain has reached 80–90% of its adult volume. "
    f"It is crucial to get these reflexes integrated early. "
    f"For these reasons, I am recommending co-treating with Stephanie Kononovich to organize and integrate these reflexes. "
    f"She will give you easy activities to do at home to speed up our results."
    ),
"retained_primitive_reflexes": lambda ctx: (
    f"Retained Primitive Reflexes – Retained primitive reflexes were observed, notably {ctx.get('primitive_reflex_findings','___')}. "
    f"This can be retained due to a birth that the baby experienced as traumatic. "
    f"These babies will cry for seemingly no reason while being required to lie back during treatments (which are painless). "
    f"Sometimes they seem terrified, even just lying down on the table — this can cause so much stress and tension in these little ones. "
    f"My cranial results can be delayed due to the baby not relaxing, not to mention the challenges of being fearful and stressed in daily life. "
    f"These reflexes can be easily calmed with appropriate exercises, and babies can then enjoy their environments without fear and stress. "
    f"Peak synaptic development is at 8 months. The cerebellum increases in size by 240% the first year of life. "
    f"By 2 years the brain has reached 80–90% of its adult volume. "
    f"It is crucial to get these reflexes integrated early. "
    f"For these reasons, I am recommending co-treating with Stephanie Kononovich to organize and integrate these reflexes. "
    f"Scheduling your appointment with her just prior to treatment with me is the best arrangement for a pleasant and successful outcome."
    ),
}

# Ordered list of findings to fill obj_fin_1..obj_fin_24
finding_order = [
    "cephalohematoma",
    "post_cephalohematoma_calcium_deposition",
    "brachycephaly",
    "plagiocephaly",
    "dolichocephaly",
    "torticollis",
    "tethered_oral_tissues",
    "internal_maxilla_rotation",
    "sutural_restrictions",
    "metopic_ridging",
    "potential_supernumerary_cranial_bones",
    "skeletal_alignment",
    "acetabular_click",
    "femur_rotation",
    "thrush",
    "cradle_cap",
    "blocked_tear_duct",
    "facial_asymmetry",
    "occipital_extension",
    "frontal_alignment",
    "temporal_alignment",
    "retrognathia",
    "prognathia",
    "inappropriate_reflexes",
    "retained_primitive_reflexes",
]

def insert_bullet_list_after_paragraph(paragraph, items):
    """
    Inserts each item as a bullet paragraph immediately after the placeholder paragraph.
    Then the caller should delete the placeholder paragraph.
    """
    current_xml = paragraph._p
    parent_container = paragraph._parent

    for item in items:
        if not str(item).strip():
            continue

        new_p = parent_container.add_paragraph(f"• {str(item).strip()}")

        current_xml.addnext(new_p._p)
        current_xml = new_p._p


def replace_placeholder_with_bullet_list(doc, placeholder, items):
    """
    Replaces a single placeholder paragraph with a dynamic bullet list.
    Works in normal document paragraphs and table cells.
    """

    # Main body paragraphs
    for p in doc.paragraphs:
        if placeholder in p.text:
            insert_bullet_list_after_paragraph(p, items)
            delete_paragraph(p)
            return True

    # Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if placeholder in p.text:
                        insert_bullet_list_after_paragraph(p, items)
                        delete_paragraph(p)
                        return True

    return False

def set_paragraph_style_safe(paragraph, *style_names):
    for style_name in style_names:
        try:
            paragraph.style = style_name
            return
        except Exception:
            pass

def is_yes(value):
    return str(value).strip().lower().startswith("y")


def is_meaningful(value):
    value = str(value).strip()
    if not value:
        return False

    ignore_values = [
        "choose an option",
        "no",
        "none",
        "no (skip the next question)",
        "yes (skip the next 2 questions)"
    ]

    return value.lower() not in ignore_values


def hours_phrase(value, label):
    """
    Prevents output like '15 hours hours'.
    If user types '15', returns '15 hours of labor'.
    If user types '15 hours', returns '15 hours of labor'.
    """
    value = str(value).strip()

    if not value:
        return ""

    if "hour" in value.lower():
        return f"{value} of {label}"

    return f"{value} hours of {label}"


def build_patient_overview(ctx):
    first = ctx.get("patient_first_name", "").strip()
    full_name = ctx.get("patient_name", "").strip()
    pronoun = ctx.get("pronoun_subj", "").strip() or "they"

    sentences = []

    # --- Birth sentence ---
    birth_details = []

    if is_yes(ctx.get("ivf_conception", "")):
        birth_details.append("was conceived by IVF")

    # Medicated / unmedicated wording
    unmedicated_value = ctx.get("unmedicated", "")

    if unmedicated_value.startswith("Yes"):
        birth_details.append("was born following an unmedicated birth")
    elif unmedicated_value == "No":
        birth_details.append("was born following a medicated birth")
    else:
        birth_details.append("was born")

    if is_yes(ctx.get("induced", "")):
        birth_details.append("with medical induction")

    if is_yes(ctx.get("epidural", "")):
        birth_details.append("with epidural anesthesia")

    birth_mode = ctx.get("birth_mode", "").strip()
    if is_meaningful(birth_mode):
        birth_details.append(f"via {birth_mode} delivery")

    weeks = ctx.get("weeks_at_birth", "").strip()
    if weeks:
        birth_details.append(f"at {weeks}")

    if birth_details:
        sentences.append(f"{first} " + ", ".join(birth_details) + ".")

    # --- Internal fetal monitor ---
    if is_yes(ctx.get("ifm", "")):
        sentences.append("An internal fetal monitor was used during labor.")

    # --- Labor / pushing ---
    labor_text = hours_phrase(ctx.get("labor_length", ""), "labor")
    pushing_text = hours_phrase(ctx.get("pushing_length", ""), "pushing")

    if labor_text and pushing_text:
        sentences.append(f"Labor included {labor_text} and {pushing_text}.")
    elif labor_text:
        sentences.append(f"Labor included {labor_text}.")
    elif pushing_text:
        sentences.append(f"Labor included {pushing_text}.")

    # --- Complications ---
    if is_yes(ctx.get("complications", "")):
        comp_text = ctx.get("complications_text", "").strip()
        if comp_text:
            sentences.append(f"Reported birth complications included {comp_text}.")
        else:
            sentences.append("Birth complications were reported, though details were not provided.")
    elif ctx.get("complications", "").strip().lower() == "no":
        sentences.append("No birth complications were reported.")

    # --- Birth weight ---
    birth_weight = ctx.get("birth_weight", "").strip()
    if birth_weight:
        sentences.append(f"{first} weighed {birth_weight} at birth.")

    # --- Breastfeeding history ---
    initial_bf = ctx.get("breastfeeding_duration_initial", "").strip()
    if initial_bf:
        sentences.append(f"{first} was initially breastfed for {initial_bf}.")

    # --- Current feeding ---
    current_feeding = ctx.get("current_feeding", "").strip()
    if current_feeding:
        sentences.append(f"{first} is currently fed by {current_feeding}.")

    # --- Breastfeeding complaints ---
    breast_complaint = ctx.get("breast_complaint", "").strip()
    if is_meaningful(breast_complaint):
        sentences.append(f"Mom reports {breast_complaint}.")

    # --- Digestive symptoms ---
    digestive = ctx.get("breast_digestive_issues", "").strip()
    if is_meaningful(digestive):
        sentences.append(f"{first} exhibits {digestive}.")

    # --- Formula ---
    formula_use = ctx.get("formula_use", "").strip()
    formula_brand = ctx.get("formula_brand", "").strip()

    if formula_use == "Primary feeding method":
        if formula_brand:
            sentences.append(f"{first} is currently receiving {formula_brand} formula.")
        else:
            sentences.append(f"{first} is currently receiving formula.")
    elif formula_use == "Supplemental only":
        if formula_brand:
            sentences.append(f"{first} is supplementally receiving {formula_brand} formula.")
        else:
            sentences.append(f"{first} is supplementally receiving formula.")

    # --- Solids ---
    solids = ctx.get("solids", "").strip()
    solids_age = ctx.get("solids_age", "").strip()

    if is_yes(solids):
        if solids_age:
            sentences.append(f"Solids were introduced at {solids_age}.")
        else:
            sentences.append("Solids have been introduced.")

    # --- Main complaint ---
    main_complaint = ctx.get("main_complaint", "").strip()
    if main_complaint:
        sentences.append(f"{full_name} was brought to the clinic with concerns of {main_complaint}.")

    # --- Previous care ---
    previous_care = ctx.get("previous_care", "").strip()
    other_care = ctx.get("other_care_details", "").strip()

    if is_yes(previous_care):
        if other_care:
            sentences.append(f"Previous care for this concern included {other_care}.")
        else:
            sentences.append("Previous care for this concern was reported.")
    elif previous_care.lower().startswith("no"):
        sentences.append("No previous care for this concern was reported.")

    return " ".join(sentences)

def replace_placeholder_with_image_grid(doc, placeholder, image_paths, captions=None, cols=3, img_width_in=2.0):
    """
    Find a placeholder paragraph in the doc (body or table cells) and replace
    it with an image grid table.  If no images are selected the placeholder
    paragraph is simply removed.
    """
    paths = [p for p in image_paths if p and os.path.exists(p)]

    def _build_and_insert(anchor_para):
        """Build image grid table and splice it after anchor_para, then delete anchor_para."""
        if not paths:
            delete_paragraph(anchor_para)
            return True

        total_w = usable_width_inches(doc)
        cell_w = total_w / cols
        img_w = min(img_width_in, max(0.8, cell_w - 0.2))
        n_rows = (len(paths) + cols - 1) // cols

        table = doc.add_table(rows=n_rows, cols=cols)
        table.autofit = False
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        idx = 0
        for r in range(n_rows):
            for c in range(cols):
                if idx >= len(paths):
                    break
                cell = table.rows[r].cells[c]
                cell.width = Inches(cell_w)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
                cp = cell.paragraphs[0]
                cp.text = ""
                cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                path = paths[idx]
                compressed = compress_image_for_report(path)
                cp.add_run().add_picture(compressed, width=Inches(img_w))
                caption = (captions or {}).get(path, "")
                if caption:
                    cap_p = cell.add_paragraph(caption)
                    cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if cap_p.runs:
                        cap_p.runs[0].italic = True
                idx += 1

            # Keep every row together — prevents image/caption split across pages
            for row in table.rows:
                prevent_row_split(row)

        # Move the table XML to sit directly after the placeholder paragraph
        anchor_para._p.addnext(table._tbl)
        delete_paragraph(anchor_para)
        return True

    # Search top-level body paragraphs first
    for p in doc.paragraphs:
        if placeholder in p.text:
            _build_and_insert(p)
            return

    # Fall back to searching inside table cells
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if placeholder in p.text:
                        _build_and_insert(p)
                        return

def add_image_grid(doc, image_paths, cols=3, img_width_in=2.0, captions=None):
    """
    Insert images into the document in a grid table.
    captions: optional dict mapping image path → caption string shown below each image.
    Patient image descriptions are looked up automatically from patient_image_descriptions.
    """
    paths = [p for p in image_paths if p and os.path.exists(p)]
    if not paths:
        return

    rows = (len(paths) + cols - 1) // cols
    total_w = usable_width_inches(doc)
    cell_w = total_w / cols
    img_w = min(img_width_in, max(0.8, cell_w - 0.2))

    table = doc.add_table(rows=rows, cols=cols)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    idx = 0
    for r in range(rows):
        for c in range(cols):
            if idx >= len(paths):
                break
            cell = table.rows[r].cells[c]
            cell.width = Inches(cell_w)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

            cp = cell.paragraphs[0]
            cp.text = ""
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER

            path = paths[idx]
            compressed = compress_image_for_report(path)
            run = cp.add_run()
            run.add_picture(compressed, width=Inches(img_w))

            # Caption: check caller-supplied captions first, then patient descriptions
            caption = (captions or {}).get(path, "")
            if not caption:
                key_match = next((k for k, v in patient_images.items() if v == path), None)
                if key_match:
                    caption = patient_image_descriptions.get(key_match, "").strip()

            if caption:
                cap_p = cell.add_paragraph(caption)
                cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cap_p.runs[0].italic = True

            idx += 1

    # Keep every row together — prevents image/caption split across pages
    for row in table.rows:
        prevent_row_split(row)

    doc.add_paragraph("")

# Main generate_report function
def generate_report():
    try:
        template_path = os.path.join(BASE_DIR, "report_template.docx")
        if not os.path.exists(template_path):
            root.after(0, lambda: messagebox.showerror(
                "Template missing", f"Template not found: {template_path}"))
            return

        doc = Document(template_path)

        # Build context dict from GUI variables (use .get() where appropriate)
        ctx = {}
        # Patient demographics
        ctx['patient_first_name'] = patient_first_name.get().strip()
        ctx['patient_last_name'] = patient_last_name.get().strip()
        ctx['patient_name'] = f"{ctx['patient_first_name']} {ctx['patient_last_name']}".strip()
        ctx['patient_dob'] = patient_dob.get().strip()
        ctx['patient_gender'] = patient_gender.get().strip()
        ctx["date_of_service"] = date_of_service.get().strip()

        # Pronouns
        pron = build_pronouns(ctx['patient_gender'])
        ctx.update(pron)

        # Birth & feeding fields (use .get() if defined)
        
        if unmedicated.get().startswith("Yes"):
            ctx["birth_med"] = "unmedicated"
        elif unmedicated.get() == "No":
            ctx["birth_med"] = "medicated"
        else:
            ctx["birth_med"] = ""
        ctx["ivf_conception"] = ivf_conception.get()
        ctx["unmedicated"] = unmedicated.get()
        ctx["induced"] = induced.get()
        ctx["epidural"] = epidural.get()
        ctx["birth_mode"] = birth_mode.get()
        ctx["weeks_at_birth"] = weeks_at_birth.get()
        ctx["ifm"] = ifm.get()
        ctx["labor_length"] = labor_length.get()
        ctx["pushing"] = pushing.get()
        ctx["pushing_length"] = pushing.get()
        ctx["complications"] = complications.get()
        ctx["complications_text"] = complications_text.get()
        ctx["birth_weight"] = birth_weight.get()
        ctx["breastfed_length"] = breastfeeding_duration_current.get()
        ctx["breastfeeding_duration_initial"] = breastfeeding_duration_initial.get()

        feeding_methods = []

        if feeding_breast.get():
            feeding_methods.append("breastfeeding")
        fu = formula_use.get().strip()
        if fu == "Primary feeding method":
            feeding_methods.append("formula")
        elif fu == "Supplemental only":
            feeding_methods.append("supplemental formula")
        if feeding_breast_bottle.get():
            feeding_methods.append("pumped breastmilk via bottle")
        if feeding_solid.get():
            feeding_methods.append("solid or pureed foods")

        ctx["feeding_breast"] = "breastfeeding" if feeding_breast.get() else ""
        ctx["feeding_formula"] = fu if fu != "None" else ""
        ctx["feeding_breast_bottle"] = "pumped breastmilk via bottle" if feeding_breast_bottle.get() else ""
        ctx["feeding_solid"] = "solid or pureed foods" if feeding_solid.get() else ""
        ctx["current_feeding"] = ", ".join(feeding_methods)

        ctx["breast_complaint"] = breast_complaint.get()
        ctx["breastfed_complaint"] = breast_complaint.get()
        ctx["breastfed_complaints"] = breast_complaint.get()
        ctx["breast_digestive_issues"] = breast_digestive_issues.get()
        ctx["patient_digestive_complaints"] = breast_digestive_issues.get()
        ctx["formula_use"] = formula_use.get()
        ctx["formula_brand"] = formula_type.get()
        ctx["formula_type"] = formula_type.get()
        ctx["solids"] = solids.get()
        ctx["solids_age"] = solids_age.get()
        ctx["previous_care"] = previous_care.get()
        ctx["previous_care_details"] = other_care_details.get()
        ctx["other_care_details"] = other_care_details.get()

        # Treatment plan and practitioner
        ctx["treatment_plan"] = treatment_plan.get().strip()

        ctx["practitioner_prefix"] = practitioner_prefix.get().strip()
        ctx["practitioner_first_name"] = practitioner_first_name.get().strip()
        ctx["practitioner_last_name"] = practitioner_last_name.get().strip()
        ctx["practitioner_suffix"] = practitioner_suffix.get().strip()

        ctx["practitioner_name"] = " ".join(
            part for part in [
                ctx["practitioner_prefix"],
                ctx["practitioner_first_name"],
                ctx["practitioner_last_name"],
                ctx["practitioner_suffix"]
            ] if part
)

        ctx["practitioner_notes"] = get_text_widget_content(Additional_Notes)

        # Recommendation text boxes
        ctx["thrush_rec"] = get_text_widget_content(thrush_rec_box)        
        ctx["cradle_cap_rec"] = get_text_widget_content(cradle_cap_rec_box)
        ctx['cephalohematoma_rec'] = get_text_widget_content(cephalohematoma_rec_box)
        ctx['blocked_tear_duct_rec'] = get_text_widget_content(blocked_tear_duct_rec_box)
        ctx['jaw_ex'] = get_text_widget_content(jaw_ex_box)
        ctx['homeo_rec'] = get_text_widget_content(homeo_rec_box)
        ctx['primitive_reflex_rec'] = get_text_widget_content(primitive_reflex_rec_box)
        ctx['homeo_cons_rec'] = get_text_widget_content(homeo_cons_rec_box)
        ctx['tots_docs_rec'] = get_text_widget_content(tots_docs_rec_box)
        ctx["primitive_reflex_findings"] = get_text_widget_content(primitive_reflexes_findings)

        ctx["patient_overview"] = build_patient_overview(ctx)


        # Additional detail fields used in templates
        detail_fields = [
            'cephalohematoma_measurement','cephalohematoma_location','cephalohematoma_cross_suture',
            'cephalohematoma_calcium','cephalohematoma_calcium_location',
            'brachycephaly_severity','brachycephaly_measurement',
            'plagiocephaly_severity','plagiocephaly_measurement','plagiocephaly_laterality',
            'dolichocephaly_severity','dolichocephaly_measurement',
            'torticollis_laterality',
            'blocked_tear_duct_laterality',
            'acetabular_click_laterality',
            'femur_rotation_laterality','femur_rotation_restriction',
            'temporal_ear_flare_laterality','temporal_alignment_outcome'
        ]
        for f in detail_fields:
            ctx[f] = globals().get(f, tk.StringVar()).get() if globals().get(f, None) is not None else ""

        # Build patient overview sentence pieces per your rules
        # IVF phrase
        ivf_phrase = "was conceived by IVF and" if ctx.get('ivf_conception','').lower().startswith('y') else ""
        # birth_med / induced / epidural phrases
        birth_med_phrase = ctx.get('birth_med','').strip()
        induced_phrase = "medically induced" if ctx.get('induced','').lower().startswith('y') else ""
        epidural_phrase = "with epidural" if ctx.get('epidural','').lower().startswith('y') else ""
        ifm_phrase = "An internal Fetal Monitor was employed." if ctx.get('ifm','').lower().startswith('y') else ""
        complications_phrase = ("complications including " + ctx.get('complications_text','').strip()) if ctx.get('complications','').lower().startswith('y') else "no complications"
        # formula
        formula_phrase = ""
        fu = ctx.get('formula_use','').strip()
        if fu == "Primary feeding method":
            formula_phrase = "is bottle-fed with " + (ctx.get('formula_brand','').strip() or "")
        elif fu == "Supplemental only":
            formula_phrase = "is supplementally bottle-fed with " + (ctx.get('formula_brand','').strip() or "")
        else:
            formula_phrase = ""
        # solids statement
        solids_statement = ""
        if ctx.get('solids','').lower().startswith('y'):
            solids_statement = f"{ctx['patient_first_name']} began taking solids at {ctx.get('solids_age','').strip()}."
        # previous care statement
        previous_care_statement = ""
        if ctx.get('previous_care','').lower().startswith('y'):
            previous_care_statement = f"received previous care for this condition with {ctx.get('previous_care_details','').strip()}."


        # Replace simple placeholders (single-line)
        simple_map = {
            "{{ date_of_service }}": ctx["date_of_service"],
            "{{ patient_name }}": ctx['patient_name'],
            "{{ patient_first_name }}": ctx['patient_first_name'],
            "{{ patient_last_name }}": ctx['patient_last_name'],
            "{{ patient_dob }}": ctx['patient_dob'],
            "{{ patient_gender }}": ctx['patient_gender'],
            "{{ treatment_plan }}": ctx['treatment_plan'],
            "{{ practitioner_name }}": ctx['practitioner_name'],
            "{{ practitioner_notes }}": ctx['practitioner_notes'],
            "{{ birth_weight }}": ctx.get('birth_weight',''),
            "{{ weeks_at_birth }}": ctx.get('weeks_at_birth',''),
            "{{ labor_length }}": ctx.get('labor_length',''),
            "{{ pushing_length }}": ctx.get('pushing_length',''),
            "{{ breastfed_length }}": ctx.get('breastfed_length',''),
            "{{ current_feeding }}": ctx.get('current_feeding',''),
            "{{ breastfed_complaints }}": ctx.get('breastfed_complaints',''),
            "{{ patient_digestive_complaints }}": ctx.get('patient_digestive_complaints',''),
            "{{ main_complaint }}": globals().get('main_complaint', tk.StringVar()).get() if globals().get('main_complaint', None) is not None else "",
            "{{ birth_mode }}": ctx.get("birth_mode", ""),
            "{{ unmedicated }}": ctx.get("unmedicated", ""),
            "{{ breastfeeding_duration_initial }}": ctx.get("breastfeeding_duration_initial", ""),
            "{{ complications_text }}": ctx.get("complications_text", ""),

            "{{ feeding_breast }}": ctx.get("feeding_breast", ""),
            "{{ feeding_formula }}": ctx.get("feeding_formula", ""),
            "{{ feeding_breast_bottle }}": ctx.get("feeding_breast_bottle", ""),
            "{{ feeding_solid }}": ctx.get("feeding_solid", ""),

            "{{ breastfed_complaint }}": ctx.get("breastfed_complaint", ""),
            "{{ breast_complaint }}": ctx.get("breast_complaint", ""),
            "{{ breast_digestive_issues }}": ctx.get("breast_digestive_issues", ""),

            "{{ formula_brand }}": ctx.get("formula_brand", ""),
            "{{ other_care_details }}": ctx.get("other_care_details", ""),
            "{{ previous_care }}": ctx.get("previous_care", ""),
            "{{ pushing }}": ctx.get("pushing", ""),
            "{{ solids }}": ctx.get("solids", ""),
            "{{ patient_overview }}": ctx.get("patient_overview", ""),
            "{{ obstacles_to_improvement }}": (
                'There are 3 "obstacles to improvement" in this process:\n'
                '1. Unreleased TOTs. Please contact one of the recommended pediatric dentists as soon as possible.\n'
                '2. Baby spending time or sleeping on their back, or spending time in restraining devices that push on the back of the head. '
                'Teach your baby to tummy sleep and keep them off the back of their head as much as possible.\n'
                '3. Retained Primitive Reflexes - tummy sleeping will help. '
                'Stephanie Kononovich can teach you how to integrate these important reflexes. '
                'https://foundationhomeopathy.com/'
            ),
            "{{ closing_paragraph }}": (
                f"It's always a pleasure to work with wonderful families like the {ctx.get('patient_last_name','')}s. "
                f"I am so glad they are under your excellent care, as well! "
                f"Please feel free to reach out anytime for discussion or questions. "
                f"I look forward to keeping you updated as we go along.\n\n"
                f"If you are unfamiliar with my approach, I use a specialized cranial technique designed specifically by me for young patients, "
                f"distinct from CranioSacral Therapy (CST). If you'd like to learn more about what I do, you can find additional details on my website, "
                f"DrKeilaDC.com, or check out a lecture I gave on plagiocephaly here https://youtu.be/dgGSUs931mU?si=6smUdXucyRxfldqV on YouTube. "
                f"I would be happy to set up a time to meet for lunch or give an educational lecture to your staff, if there is interest."
            ),
        }
        # Merge simple_map and overview_map into replacements
        replacements = {**simple_map}
        # Add pronouns
        replacements.update({
            "{{ pronoun_subj }}": ctx.get('pronoun_subj',''),
            "{{ pronoun_subj_cap }}": ctx.get('pronoun_subj_cap',''),
            "{{ pronoun_obj }}": ctx.get('pronoun_obj',''),
            "{{ pronoun_obj_cap }}": ctx.get('pronoun_obj_cap',''),
            "{{ pronoun_poss }}": ctx.get('pronoun_poss',''),
            "{{ pronoun_poss_cap }}": ctx.get('pronoun_poss_cap',''),
            "{{ practitioner_prefix }}": ctx.get("practitioner_prefix",""),
            "{{ practitioner_first_name }}": ctx.get("practitioner_first_name",""),
            "{{ practitioner_last_name }}": ctx.get("practitioner_last_name",""),
            "{{ practitioner_suffix }}": ctx.get("practitioner_suffix",""),
            })

        # Replace all simple placeholders in doc
        for ph, val in replacements.items():
            replace_placeholder(doc, ph, val)

        # ---------- BULLET-FORMATTED RECOMMENDATIONS ----------

        recommendation_fields = [
            ("{{ thrush_rec }}", ctx["thrush_rec"]),
            ("{{ cradle_cap_rec }}", ctx["cradle_cap_rec"]),
            ("{{ cephalohematoma_rec }}", ctx["cephalohematoma_rec"]),
            ("{{ blocked_tear_duct_rec }}", ctx["blocked_tear_duct_rec"]),
            ("{{ jaw_ex }}", ctx["jaw_ex"]),
            ("{{ homeo_rec }}", ctx["homeo_rec"]),
            ("{{ primitive_reflex_rec }}", ctx["primitive_reflex_rec"]),
            ("{{ homeo_cons_rec }}", ctx["homeo_cons_rec"]),
            ("{{ tots_docs_rec }}", ctx["tots_docs_rec"]),

        ]

        recommendation_items = []

        for placeholder, text in recommendation_fields:
            if text.strip():
                recommendation_items.append(text.strip())

        replace_placeholder_with_bullet_list(
            doc,
            "{{ recommendations_list }}",
            recommendation_items
)


        # Build objective findings paragraphs list based on toggles
        obj_paragraphs = []
        for key in finding_order:
            include = False
            # Determine include logic: check boolean variables by name
            # Map finding key to GUI boolean variable names where applicable
            mapping_bool = {
                "cephalohematoma": 'cephalohematoma',
                "post_cephalohematoma_calcium_deposition": 'cephalohematoma',  # include only if cephalohematoma exists and calcium present
                "brachycephaly": 'brachycephaly',
                "plagiocephaly": 'plagiocephaly',
                "dolichocephaly": 'dolichocephaly',
                "torticollis": 'torticollis',
                "tethered_oral_tissues": 'tethered_oral_tissues',
                "internal_maxilla_rotation": 'internal_maxilla_rotation',
                "sutural_restrictions": 'sutural_restriction',  # if you have a toggle for this, else include based on other cues
                "metopic_ridging": 'metopic_ridge',
                "potential_supernumerary_cranial_bones": 'supernumerary_bones',
                "skeletal_alignment": 'skeletal_alignment',
                "acetabular_click": 'acetabular_click',
                "femur_rotation": 'femur_rotation',
                "thrush": 'thrush',
                "cradle_cap": 'cradle_cap',
                "blocked_tear_duct": 'blocked_tear_duct',
                "facial_asymmetry": 'facial_asymmetry',
                "occipital_extension": 'occipital_extension',
                "frontal_alignment": 'frontal_alignment',
                "temporal_alignment": 'temporal_alignment',
                "retrognathia": 'retrognathia',
                "prognathia": 'prognathia',
                "inappropriate_reflexes": 'inappropriate_reflexes',
                #"retained_primitive_reflexes": 'retained_primitive_reflexes',
            }
            bool_var_name = mapping_bool.get(key)
            if bool_var_name:
                gv = globals().get(bool_var_name)
                if isinstance(gv, tk.BooleanVar):
                    include = bool(gv.get())
                else:
                    # if no BooleanVar exists, fallback to checking presence of detail fields
                    include = any(ctx.get(fld) for fld in detail_fields)
            # Special case: post_cephalohematoma_calcium_deposition only if cephalohematoma and calcium == Yes
            if key == "post_cephalohematoma_calcium_deposition":
                include = False
                if globals().get('cephalohematoma') and cephalohematoma.get():
                    if ctx.get('cephalohematoma_calcium','').lower().startswith('y'):
                        include = True

            if include and key in finding_templates:
                text = finding_templates[key](ctx)
                obj_paragraphs.append(text)

        # ---------- BULLET-FORMATTED OBJECTIVE FINDINGS ----------
        replace_placeholder_with_bullet_list(
            doc,
            "{{ objective_findings_list }}",
            obj_paragraphs
        )


        # Insert the 5 key patient photos right after objective findings
        patient_row_keys = ["anterior_view", "superior_view", "left_profile", "right_profile", "lip_frenulum", "tongue_frenulum"]
        row_paths = [patient_images.get(k) for k in patient_row_keys]

        replace_placeholder_with_image_row(
            doc,
            "{{ patient_photo_row }}",
            row_paths,
            max_cols=6
        )

        # ---------------------------------------------------------
        # REFERENCE IMAGES — inserted at {{ reference_images }}
        # in the template (between Objective Findings and Recommendations)
        # Must run BEFORE the leftover placeholder check.
        # ---------------------------------------------------------
        reference_paths = []

        if img_right_plagio.get(): reference_paths.append(IMAGE_MAP["right_plagio"])
        if img_left_plagio.get(): reference_paths.append(IMAGE_MAP["left_plagio"])
        if img_temporal_rotation.get(): reference_paths.append(IMAGE_MAP["temporal_rotation"])
        if img_pscb_1.get(): reference_paths.append(IMAGE_MAP["supernumerary_bones_1"])
        if img_pscb_2.get(): reference_paths.append(IMAGE_MAP["supernumerary_bones_2"])
        if img_parietal_bone.get(): reference_paths.append(IMAGE_MAP["parietal_bone"])
        if img_sagittal_suture.get(): reference_paths.append(IMAGE_MAP["sagittal_suture"])
        if img_normal_lip_frenulum.get(): reference_paths.append(IMAGE_MAP["normal_lip_frenulum"])
        if img_normal_tongue_frenulum_1.get(): reference_paths.append(IMAGE_MAP["normal_tongue_frenulum_1"])
        if img_normal_tongue_frenulum_2.get(): reference_paths.append(IMAGE_MAP["normal_tongue_frenulum_2"])
        if img_eop.get(): reference_paths.append(IMAGE_MAP["eop"])

        # Build path → caption mapping for reference images
        ref_captions = {IMAGE_MAP[key]: IMAGE_TITLES[key] for key in IMAGE_TITLES if key in IMAGE_MAP}

        # Replace the {{ reference_images }} placeholder with the image grid
        replace_placeholder_with_image_grid(doc, "{{ reference_images }}", reference_paths, ref_captions, cols=4, img_width_in=1.0)

        # Check for any remaining unfilled placeholders (thread-safe warning)
        full_text = "\n".join([p.text for p in doc.paragraphs])
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    full_text += "\n" + "\n".join([p.text for p in cell.paragraphs])

        leftover = sorted(set(re.findall(r"\{\{\s*([^}]+?)\s*\}\}", full_text)))
        if leftover:
            msg = "Unfilled placeholders detected:\n\n" + "\n".join(leftover)
            root.after(0, lambda m=msg: messagebox.showwarning("Unfilled placeholders", m))

        # ---------------------------------------------------------
        # ADDITIONAL PATIENT IMAGES APPENDIX
        # ---------------------------------------------------------
        exclude_keys = {"anterior_view", "superior_view", "left_profile", "right_profile", "lip_frenulum", "tongue_frenulum"}

        other_patient_paths = [
            path for key, path in patient_images.items()
            if key not in exclude_keys and path and os.path.exists(path)
        ]

        if other_patient_paths:
            doc.add_page_break()
            doc.add_heading("Additional Patient Images", level=1)
            add_image_grid(doc, other_patient_paths, cols=3, img_width_in=2.0)


        # Save file
        safe_last = ctx['patient_last_name'] or "Patient"
        date_str = datetime.date.today().strftime("%Y%m%d")
        out_name = os.path.join(BASE_DIR, f"{safe_last}_{date_str}_Report.docx")
        doc.save(out_name)

        root.after(0, lambda: messagebox.showinfo(
            "Report saved", f"Report saved as:\n{out_name}"))

    except Exception as e:
        err = str(e)
        root.after(0, lambda: messagebox.showerror(
            "Error generating report", err))

# Add Generate Report button and status label to Practitioner Notes tab
generate_status = tk.StringVar(value="")
tk.Label(tab_practitioner, textvariable=generate_status, font=("Arial", 10), fg="#888").pack(anchor="e", padx=20)

def start_generate_report():
    """Disable the button, show status, then run generation in a background thread."""
    generate_btn.config(state="disabled", text="Generating…")
    generate_status.set("Please wait — building your report…")

    def _run():
        try:
            generate_report()
        finally:
            # Always re-enable the button on the main thread when done
            root.after(0, _generation_done)

    threading.Thread(target=_run, daemon=True).start()

def _generation_done():
    generate_btn.config(state="normal", text="Generate Report")
    generate_status.set("")

generate_btn = tk.Button(tab_practitioner, text="Generate Report", command=start_generate_report, bg="#4CAF50", fg="white")
generate_btn.pack(anchor="e", padx=20, pady=20)


# ---------- End: Generate Report Functionality ----------






bind_mousewheel_to_patient()
root.mainloop()
