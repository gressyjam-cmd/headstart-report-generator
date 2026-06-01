from importlib.resources import path
from pdb import run
from pydoc import doc
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinterdnd2 import TkinterDnD
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
from PIL import Image, ImageTk
import datetime
import os
import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT

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
import tkinter as tk
...

IMAGE_MAP = {
    "right_plagio": "images/right_plagio.jpg",
    "left_plagio": "images/left_plagio.jpg",
    "temporal_rotation": "images/temporal_rotation.jpg",
    "supernumerary_bones_1": "images/PSCB_1.jpg",
    "supernumerary_bones_2": "images/PSCB_2.jpg",
    "parietal_bone": "images/parietal_bone.jpg",
    "sagittal_suture": "images/sagittal_suture.jpg",
    "normal_lip_frenulum": "images/normal_lip_frenulum.jpg",
    "normal_tongue_frenulum": "images/normal_tongue_frenulum.jpg",
    "normal_tongue_frenulum_1": "images/normal_tongue_frenulum_1.jpg",
    "normal_tongue_frenulum_2": "images/normal_tongue_frenulum_2.jpg",
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
    "cephalohematoma": None,
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

canvas_pr = tk.Canvas(frame_practitioner)
canvas_pr.pack(side="left", fill="both", expand=True)

scroll_pr = ttk.Scrollbar(frame_practitioner, orient="vertical", command=canvas_pr.yview)
scroll_pr.pack(side="right", fill="y")

canvas_pr.configure(yscrollcommand=scroll_pr.set)
canvas_pr.bind("<Configure>", lambda e: canvas_pr.configure(scrollregion=canvas_pr.bbox("all")))

practitioner_form = tk.Frame(canvas_pr)
practitioner_window = canvas_pr.create_window((0, 0), window=practitioner_form, anchor="nw")

def resize_practitioner(event):
    canvas_pr.itemconfig(practitioner_window, width=event.width)

canvas_pr.bind("<Configure>", resize_practitioner)

# Practitioner tabs
notebook = ttk.Notebook(practitioner_form)
notebook.pack(fill="both", expand=True, padx=20, pady=20)

# Tab Frames
tab_cranial = tk.Frame(notebook)
tab_oral_neck = tk.Frame(notebook)
tab_other = tk.Frame(notebook)
tab_recommend = tk.Frame(notebook)
tab_images = tk.Frame(notebook)
tab_practitioner = tk.Frame(notebook)


for tab in (tab_cranial, tab_oral_neck, tab_other, tab_recommend, tab_practitioner):
    tab.pack(fill="both", expand=True)


#add tabs to notebook
notebook.add(tab_cranial, text="Cranial Findings")
notebook.add(tab_oral_neck, text="Oral & Neck")
notebook.add(tab_other, text="Other Findings")
notebook.add(tab_recommend, text="Recommendations")
notebook.add(tab_images, text="Images")
notebook.add(tab_practitioner, text="Practitioner Notes")

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


    return inner

EMU_PER_INCH = 914400

def usable_width_inches(doc):
    s = doc.sections[0]
    usable = s.page_width - s.left_margin - s.right_margin
    return usable / EMU_PER_INCH

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)

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




cranial_frame = make_scrollable(tab_cranial)
oral_neck_frame = make_scrollable(tab_oral_neck)
other_frame = make_scrollable(tab_other)
recommend_frame = make_scrollable(tab_recommend)
practitioner_frame = make_scrollable(tab_practitioner)
images_frame = make_scrollable(tab_images)


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

def toggle_plagiocephaly_fields():
    state = "normal" if plagiocephaly.get() else "disabled"
    plagiocephaly_severity_dropdown.config(state=state)
    plagiocephaly_measurement_entry.config(state=state)
    plagiocephaly_laterality_dropdown.config(state=state)

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

supernumerary_bones = tk.BooleanVar()
tk.Checkbutton(
    cranial_frame,
    text="Supernumerary Cranial Bones",
    variable=supernumerary_bones
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
            "Borax 30c twice daily; SFI Lacto Prime or Baby Probiotic; "
            "Ther-biotic Complete for mom; reduce sugar/simple carbs; "
            "colloidal silver on nipples; Silverette nipple shields."
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
            "Calc Sulph 6x 3x/day; SFI Lacto Prima or Baby Probiotic; "
            "massage coconut/olive oil into scalp (do not scrub)."
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
        blocked_tear_duct_rec_box.insert("1.0", "Silicea 6x 2–3x daily.")
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


##=====================
# Images Tab
##=====================

from docx.shared import Inches  # already present in your file

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
img_normal_tongue_frenulum = tk.BooleanVar()
img_normal_tongue_frenulum_1 = tk.BooleanVar()
img_normal_tongue_frenulum_2 = tk.BooleanVar()

tk.Label(images_frame, text="Clinic Reference Images", font=("Arial", 14, "bold")).pack(anchor="w", padx=20, pady=(10,5))

tk.Checkbutton(images_frame, text="Right Plagiocephaly", variable=img_right_plagio).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Left Plagiocephaly", variable=img_left_plagio).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Temporal Bone Rotation / Ear Flare", variable=img_temporal_rotation).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Supernumerary Cranial Bones (PSCB 1)", variable=img_pscb_1).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Supernumerary Cranial Bones (PSCB 2)", variable=img_pscb_2).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Parietal Bone (Cephalohematoma)", variable=img_parietal_bone).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Sagittal Suture (Cephalohematoma)", variable=img_sagittal_suture).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Normal Lip Frenulum", variable=img_normal_lip_frenulum).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Normal Tongue Frenulum", variable=img_normal_tongue_frenulum).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Normal Tongue Frenulum (Alt 1)", variable=img_normal_tongue_frenulum_1).pack(anchor="w", padx=40)
tk.Checkbutton(images_frame, text="Normal Tongue Frenulum (Alt 2)", variable=img_normal_tongue_frenulum_2).pack(anchor="w", padx=40)

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
        path = event.data.strip("{}")
        patient_images[key] = path

        # Update label
        label.config(text=f"Loaded: {path.split('/')[-1]}")

        # Load thumbnail
        try:
            img = Image.open(path)
            img.thumbnail((200, 200))
            tk_img = ImageTk.PhotoImage(img)

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


practitioner_form.bind(
    "<Configure>",
    lambda e: canvas_pr.configure(scrollregion=canvas_pr.bbox("all")))

#enable mousewheel scrolling
def _on_mousewheel_p(event):
    canvas_p.yview_scroll(int(-1 * (event.delta / 120)), "units")


def _on_mousewheel_pr(event):
    canvas_pr.yview_scroll(int(-1 * (event.delta / 120)), "units")

def bind_mousewheel_to_patient():
    root.unbind_all("<MouseWheel>")
    root.bind_all("<MouseWheel>", _on_mousewheel_p)

def bind_mousewheel_to_practitioner():
    root.unbind_all("<MouseWheel>")
    root.bind_all("<MouseWheel>", _on_mousewheel_pr)


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
feeding_formula = tk.BooleanVar()
feeding_breast_bottle = tk.BooleanVar()
feeding_solid = tk.BooleanVar()

tk.Checkbutton(form_frame, text="Breastfeeding", variable=feeding_breast).pack(anchor="w", padx=40)
tk.Checkbutton(form_frame, text="Formula", variable=feeding_formula).pack(anchor="w", padx=40)
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
tk.Label(form_frame, text="Are you currently feeding with formula?", font=("Arial", 12)).pack(anchor="w", padx=20, pady=(20, 0))
formula_use = tk.StringVar()
formula_use_dropdown = ttk.Combobox(form_frame, textvariable=formula_use, values=["Choose An Option", "Yes", "No (Skip the next 2 questions)", "Supplemental only"], width=37)
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
    print("Feeding Method - Formula:", feeding_formula.get())
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

# --- Bottom Submit Area for Practitioner Page ---
practitioner_bottom = tk.Frame(practitioner_form)
practitioner_bottom.pack(fill="x", pady=30)

tk.Button(practitioner_bottom, text="Submit All Data", command=submit_data).pack()
tk.Button(practitioner_bottom, text="Back to Patient Mode", command=show_patient_frame).pack(pady=10)




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

# Objective finding templates (edit as needed)
finding_templates = {
"cephalohematoma": lambda ctx: (
    f"Cephalohematoma – {ctx['patient_name']}'s cephalohematoma measured "
    f"{ctx.get('cephalohematoma_measurement','___')} and was located at "
    f"{ctx.get('cephalohematoma_location','___')}. "
    f"{'It crossed the sutural lines' if is_yes(ctx.get('cephalohematoma_cross_suture','')) else 'It did not cross the sutural lines'} "
    f"and has calcium deposits {ctx.get('cephalohematoma_calcium_location','___')}. "
    f"A cephalohematoma is often due to birth trauma such as prolonged transition or repeated descent. "
    f"These swellings often resolve within days, but delayed resolution indicates inefficient CSF routing. "
    f"Cranial treatments combined with homeopathy often give excellent results. "
    f"Stagnation in the sagittal sinus along the motor homunculus can contribute to seizures, which cranial work can resolve. "
    f"Occasionally, bits of blood appear in the urine as the CH resolves. Homeopathic remedies help resolve calcium deposition."
    ),
"post_cephalohematoma_calcium_deposition": lambda ctx: (
    f"Post Cephalohematoma Calcium Deposition – {ctx['patient_name']}'s cephalohematoma had resolved at the time of "
    f"{ctx['pronoun_poss']} exam, but calcium deposits are palpable along the previous CH edges. "
    f"A cephalohematoma is often due to birth trauma such as prolonged transition or repeated descent. "
    f"These swellings often resolve within days, but delayed resolution indicates inefficient CSF routing. "
    f"Cranial treatments combined with homeopathy often give excellent results. "
    f"Stagnation in the sagittal sinus along the motor homunculus can contribute to seizures. "
    f"It is important to treat after EXTERNAL resolution to ensure INTERNAL CSF flow is restored. "
    f"Homeopathic remedies help safely resolve calcium deposition."
    ),
"brachycephaly": lambda ctx: (
    f"Brachycephaly – {ctx['patient_name']} presents with a {ctx.get('brachycephaly_severity','___')} "
    f"non-synostotic brachycephaly. Non-synostotic means no sutural fusion or aplasia—"
    f"the patient is not a surgical candidate. Brachycephaly means the back of the head is flat and "
    f"the front-to-back distance is shortened relative to width. CI measurement: "
    f"{ctx.get('brachycephaly_measurement','___')}. Normal CI is closer to 80. "
    f"The occiput houses the cerebellum, which coordinates tone, balance, posture, attention, language, "
    f"and emotional regulation. A flat occiput increases scoliosis risk due to compensatory cervical curve changes."
),
"plagiocephaly": lambda ctx: (
    f"Plagiocephaly – {ctx['patient_name']} presents with {ctx.get('plagiocephaly_severity','___')} "
    f"non-synostotic plagiocephaly, with a CVA of {ctx.get('plagiocephaly_measurement','___')}. "
    f"Non-synostotic means no sutural fusion or aplasia—no surgical indication. "
    f"Plagiocephaly means one side of the head is anterior relative to the other. "
    f"Laterality: {ctx.get('plagiocephaly_laterality','___')}. "
    f"When torticollis coexists, the neck preference reinforces the cranial asymmetry, creating a cycle."
    ),
"dolichocephaly": lambda ctx: (
    f"Dolichocephaly – {ctx['patient_name']} presents with a {ctx.get('dolichocephaly_severity','___')} "
    f"non-synostotic dolichocephalic presentation. CI: {ctx.get('dolichocephaly_measurement','___')}. "
    f"Dolichocephaly means the head is longer and narrower than ideal. "
    f"A narrow cranium causes a high palate (affecting nursing) and increases risk of dental crowding. "
    f"A high palate is especially problematic with tongue ties, leading to incomplete milk extraction and low supply."
    ),
"torticollis": lambda ctx: (
    f"Torticollis – A mild to moderate {ctx.get('torticollis_laterality','___')} lateral torticollis was noted. "
    f"This is essentially a crick in the neck, causing a preference for turning to one side. "
    f"It may be due to internal dural torsion, vertebral fixation, or muscular imbalance, and can make nursing on one breast "
    f"more comfortable than the other. Home exercises such as gentle inversion swinging a couple of times daily help straighten "
    f"the dural torsion and support treatment. A good lactation consultant can help identify comfortable nursing positions. "
    f"Vagal nerve compression often accompanies torticollis and may contribute to digestive issues; this typically resolves as "
    f"the torticollis improves."
    ),
"tethered_oral_tissues": lambda ctx: (
    f"Tethered Oral Tissues (TOTs): {ctx['patient_first_name']}'s oral ties need to be evaluated for possible "
    f"laser frenectomy by a pediatric dentist extensively trained in TOTs. "
    f"{ctx['pronoun_poss_cap']} tongue looks and feels tight, and {ctx['pronoun_subj']} gags when the palate is touched, "
    f"which can indicate a restricted tongue. Oral ties can cause nursing issues, breast pain, incomplete milk extraction, "
    f"low weight gain, painful gas, diminishing milk supply, mastitis, and more. Speech, dentition, neck/throat tension, "
    f"posture, gag reflex, and airway function can all be negatively affected by not releasing ties during infancy. "
    f"TOTs can cause delayed response to cranial treatments. Infancy is the ideal time for release."
    ),
"internal_maxilla_rotation": lambda ctx: (
    f"Internal Maxilla Rotation – A high, narrow palate with a low intermaxillary suture was observed. "
    f"This is commonly associated with internally rotated maxillae and intermaxillary sutural restriction, often related to "
    f"tongue‑tie. The tongue’s pressure on the palate helps shape the maxilla; when the tongue cannot fully contact the palate, "
    f"the palate becomes high and narrow. Internal maxilla rotation is frequently associated with dark circles under the eyes."
    ),
"sutural_restrictions": lambda ctx: (
    f"Sutural Restrictions – Cranial findings present with tight vault, oral, and facial sutures with ridging along several "
    f"cranial vault sutures. This is common in C‑sections, babies trapped in transition, or those with very rapid transitions. "
    f"Sutural restriction limits cranial expansion. It may affect only one region (causing asymmetry), the sinus or oral "
    f"structures, or the entire cranial vault. Because brain growth drives cranial expansion, unrestricted sutures are essential. "
    f"Asymmetries may not appear until the 12‑week growth spurt. These findings are not merely aesthetic — cranial shape affects "
    f"brain development, TMJ function, airway development, and future orthodontic needs."
    ),
"metopic_ridging": lambda ctx: (
    f"Metopic Ridging – The metopic suture runs vertically down the center of the forehead and is the only cranial suture that "
    f"normally fuses. Fusion typically occurs between 3–9 months of age. Cranial therapy before this window can improve ridging "
    f"when the suture is jammed. After fusion, cranial therapy may still reduce the severity of ridging over growth spurts, "
    f"though progress may take months to years."
    ),
"potential_supernumerary_cranial_bones": lambda ctx: (
    f"Potential Supernumerary Cranial Bones – Small accessory bones were noted along the lambdoid region. These are normal "
    f"variants and typically of no functional or aesthetic consequence unless accompanied by flattening and sutural restriction. "
    f"Cranial work can usually restore normal shape and positioning of the occiput and parietals, though the small accessory "
    f"bones may remain slightly recessed. These variants rarely warrant imaging, as they are clinically insignificant."
    ),
"acetabular_click": lambda ctx: (
    f"Acetabular Click – An acetabular click was noticed on the "
    f"{ctx.get('acetabular_click_laterality','___')} hip. "
    f"Acetabular clicks may indicate mild hip instability or shallow acetabular development. "
    f"Evaluation by a pediatric orthopedist or imaging may be recommended depending on clinical presentation."
    ),
"femur_rotation": lambda ctx: (
    f"Femur Rotation – Decreased {ctx.get('femur_rotation_restriction','___')} rotation of the "
    f"{ctx.get('femur_rotation_laterality','___')} femur was noted. "
    f"This may reflect muscular imbalance, pelvic torsion, or intrauterine positioning patterns. "
    f"Cranial and pelvic balancing often improve femoral rotation symmetry."
    ),
"thrush": lambda ctx: (
    f"Thrush – Thrush was visibly apparent on the tongue, which can cause painful nursing for both the infant and the mother, "
    f"as well as bloating, reflux, and digestive discomfort. Homeopathic support such as Borax 30c twice daily is often helpful. "
    f"A high‑quality probiotic (SFI Lacto Prime or SFI Baby Probiotic) is recommended twice daily. "
    f"If breastfeeding, SFI Ther‑biotic Complete Probiotics may support maternal flora. "
    f"Nipples may be sprayed with colloidal silver before and after nursing. "
    f"If symptoms do not resolve within two weeks, a consultation with a homeopath is recommended."
    ),
"cradle_cap": lambda ctx: (
    f"Cradle Cap – Cradle cap was noted. This is a fungal condition often associated with a temporarily weakened immune system. "
    f"Parents should avoid scrubbing the flakes, as this can worsen the condition. "
    f"Homeopathic and probiotic support may be beneficial. If no improvement is seen within 3–4 weeks, "
    f"a consultation with a homeopath is recommended for a more specific remedy."
    ),
"blocked_tear_duct": lambda ctx: (
    f"Blocked Tear Duct – A blocked {ctx.get('blocked_tear_duct_laterality','___')} tear duct was noted. "
    f"The homeopathic cell salt Silicea 6x is often recommended twice daily. "
    f"Cranial and sutural release may also support improved drainage."
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
    f"anterior malar positioning, or chin deviation. Although sometimes considered aesthetic, frontal alignment affects TMJ function, "
    f"vision, and maxillary rotation."
    ),
"temporal_alignment": lambda ctx: (
    f"Temporal Alignment – A {ctx.get('temporal_ear_flare_laterality','___')} ear flare was noted, typically due to external rotation "
    f"of one or both temporal bones. Uncorrected temporal rotation can contribute to crossbite tendencies, recessed jaw posture, "
    f"sensitive ears, and motion sickness."
    ),
"retrognathia": lambda ctx: (
    f"Retrognathia – A recessed chin was noted, often due to externally rotated temporal bones affecting the TMJ and ear canal. "
    f"This can contribute to nursing discomfort and may predispose to orthodontic challenges later in life. "
    f"Jaw exercises have been recommended."
    ),
"prognathia": lambda ctx: (
    f"Prognathia – An underbite tendency was noted. This may result from internal rotation of the temporal bones, mouth breathing, "
    f"or prolonged thumb‑sucking. Cranial balancing and oral‑motor support may help improve jaw alignment."
    ),
"inappropriate_reflexes": lambda ctx: (
    f"Inappropriately Exhibited Primitive Reflexes – Primitive reflexes were observed including: "
    f"{ctx.get('primitive_reflex_findings','___')}. "
    f"Disorganized reflexes can delay treatment results and contribute to daily challenges. "
    f"These reflexes can be calmed and integrated with specific exercises. "
    f"Early integration is important due to rapid cerebellar and synaptic development in infancy."
    ),
"retained_primitive_reflexes": lambda ctx: (
    f"Retained Primitive Reflexes – Retained primitive reflexes were observed, notably "
    f"{ctx.get('primitive_reflex_findings','___')}. "
    f"These may persist after a birth the infant experienced as traumatic. "
    f"Babies may cry or appear fearful when lying back for treatment. "
    f"Retained reflexes can cause tension, stress, and delayed cranial results. "
    f"Early integration is crucial due to rapid brain growth in the first two years of life."
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

def delete_paragraph(paragraph):
    p = paragraph._element
    parent = p.getparent()
    parent.remove(p)


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
    pushing_text = hours_phrase(ctx.get("pushing", ""), "pushing")

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

    if formula_use == "Yes":
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

def add_image_grid(doc, image_paths, cols=3, img_width_in=2.0):
    paths = [p for p in image_paths if p and os.path.exists(p)]
    if not paths:
        return

    rows = (len(paths) + cols - 1) // cols
    total_w = usable_width_inches(doc)
    cell_w = total_w / cols
    # keep images slightly smaller than cell
    img_w = min(img_width_in, max(0.8, cell_w - 0.2))

    table = doc.add_table(rows=rows, cols=cols)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    idx = 0
    for r in range(rows):
        for c in range(cols):
            cell = table.rows[r].cells[c]
            cell.width = Inches(cell_w)  # reliable sizing in Word [1](https://pytutorial.com/python-docx-paragraph-formatting-guide/)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

            cp = cell.paragraphs[0]      # use existing paragraph [2](https://skelmis-docx.readthedocs.io/en/stable/user/styles-using.html)
            cp.text = ""
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER

        path = paths[idx]
        run = cp.add_run()
        run.add_picture(path, width=Inches(img_w))

        # --- Add description BELOW image ---
        key_match = next((k for k, v in patient_images.items() if v == path), None)

        if key_match:
            desc = patient_image_descriptions.get(key_match, "").strip()
            if desc:
                desc_p = cell.add_paragraph(desc)
                desc_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        idx += 1

    doc.add_paragraph("")

# Main generate_report function
def generate_report():
    try:
        template_path = "report_template.docx"
        if not os.path.exists(template_path):
            messagebox.showerror("Template missing", f"Template not found: {template_path}")
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
        ctx["formula_brand"] = formula_type.get()
        ctx["previous_care_details"] = other_care_details.get()
        ctx["breastfed_length"] = breastfeeding_duration_current.get()
        ctx["ivf_conception"] = ivf_conception.get()
        ctx["unmedicated"] = unmedicated.get()
        ctx["induced"] = induced.get()
        ctx["epidural"] = epidural.get()
        ctx["birth_mode"] = birth_mode.get()
        ctx["weeks_at_birth"] = weeks_at_birth.get()
        ctx["ifm"] = ifm.get()
        ctx["labor_length"] = labor_length.get()
        ctx["pushing_length"] = pushing.get()
        ctx["complications"] = complications.get()
        ctx["complications_text"] = complications_text.get()
        ctx["birth_weight"] = birth_weight.get()
        ctx["breastfed_length"] = breastfeeding_duration_current.get()

        feeding_methods = []

        if feeding_breast.get():
            feeding_methods.append("breastfeeding")
        if feeding_formula.get():
            feeding_methods.append("formula")
        if feeding_breast_bottle.get():
            feeding_methods.append("pumped breastmilk via bottle")
        if feeding_solid.get():
            feeding_methods.append("solid or pureed foods")

        # Individual feeding placeholders, if your template ever uses them
        ctx["feeding_breast"] = "breastfeeding" if feeding_breast.get() else ""
        ctx["feeding_formula"] = "formula" if feeding_formula.get() else ""
        ctx["feeding_breast_bottle"] = "pumped breastmilk via bottle" if feeding_breast_bottle.get() else ""
        ctx["feeding_solid"] = "solid or pureed foods" if feeding_solid.get() else ""

        # Combined feeding sentence placeholder
        ctx["current_feeding"] = ", ".join(feeding_methods)
        
        # Template-compatible aliases for patient overview
        ctx["breastfeeding_duration_initial"] = breastfeeding_duration_initial.get()
        ctx["breast_complaint"] = breast_complaint.get()
        ctx["breastfed_complaint"] = breast_complaint.get()
        ctx["breastfed_complaints"] = breast_complaint.get()

        ctx["breast_digestive_issues"] = breast_digestive_issues.get()
        ctx["patient_digestive_complaints"] = breast_digestive_issues.get()

        ctx["formula_brand"] = formula_type.get()
        ctx["formula_type"] = formula_type.get()

        ctx["pushing"] = pushing.get()
        ctx["pushing_length"] = pushing.get()

        ctx["previous_care"] = previous_care.get()
        ctx["other_care_details"] = other_care_details.get()

        ctx["solids"] = solids.get()
        ctx["solids_age"] = solids_age.get()

        ctx["breastfed_complaints"] = breast_complaint.get()
        ctx["patient_digestive_complaints"] = breast_digestive_issues.get()
        ctx["formula_use"] = formula_use.get()
        ctx["formula_type"] = formula_type.get()
        ctx["solids"] = solids.get()
        ctx["solids_age"] = solids_age.get()
        ctx["previous_care"] = previous_care.get()
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
        fu = ctx.get('formula_use','').strip().lower()
        if fu == "yes" or fu == "y":
            formula_phrase = "is bottle-fed with " + (ctx.get('formula_brand','').strip() or "")
        elif fu.startswith("supplement"):
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

        # after all replacements and before doc.save(...)
        full_text = "\n".join([p.text for p in doc.paragraphs])
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    full_text += "\n" + "\n".join([p.text for p in cell.paragraphs])

        leftover = sorted(set(re.findall(r"\{\{\s*([^}]+?)\s*\}\}", full_text)))
        if leftover:
            msg = "Unfilled placeholders detected:\n\n" + "\n".join(leftover)
            messagebox.showwarning("Unfilled placeholders", msg)

        # ---------------------------------------------------------
        # IMAGE APPENDIX (COMPACT GRID)
        # ---------------------------------------------------------

        doc.add_page_break()
        doc.add_heading("Image Appendix", level=1)

        # -------------------------
        # Reference images (checkbox-driven)
        # -------------------------
        reference_paths = []

        if img_right_plagio.get(): reference_paths.append(IMAGE_MAP["right_plagio"])
        if img_left_plagio.get(): reference_paths.append(IMAGE_MAP["left_plagio"])
        if img_temporal_rotation.get(): reference_paths.append(IMAGE_MAP["temporal_rotation"])
        if img_pscb_1.get(): reference_paths.append(IMAGE_MAP["supernumerary_bones_1"])
        if img_pscb_2.get(): reference_paths.append(IMAGE_MAP["supernumerary_bones_2"])
        if img_parietal_bone.get(): reference_paths.append(IMAGE_MAP["parietal_bone"])
        if img_sagittal_suture.get(): reference_paths.append(IMAGE_MAP["sagittal_suture"])
        if img_normal_lip_frenulum.get(): reference_paths.append(IMAGE_MAP["normal_lip_frenulum"])
        if img_normal_tongue_frenulum.get(): reference_paths.append(IMAGE_MAP["normal_tongue_frenulum"])
        if img_normal_tongue_frenulum_1.get(): reference_paths.append(IMAGE_MAP["normal_tongue_frenulum_1"])
        if img_normal_tongue_frenulum_2.get(): reference_paths.append(IMAGE_MAP["normal_tongue_frenulum_2"])

        doc.add_heading("Reference Images", level=2)
        add_image_grid(doc, reference_paths, cols=3, img_width_in=2.0)


    # -------------------------
        # Additional patient images (everything EXCEPT the 5 main photos)
        # -------------------------
        exclude_keys = {"anterior_view", "superior_view", "left_profile", "right_profile", "lip_frenulum", "tongue_frenulum"}

        other_patient_paths = [
            path for key, path in patient_images.items()
            if key not in exclude_keys and path and os.path.exists(path)
        ]

        if other_patient_paths:
            doc.add_heading("Additional Patient Images", level=2)
            add_image_grid(doc, other_patient_paths, cols=3, img_width_in=2.0)


        # Save file
        safe_last = ctx['patient_last_name'] or "Patient"
        date_str = datetime.date.today().strftime("%Y%m%d")
        out_name = f"{safe_last}_{date_str}_Report.docx"
        doc.save(out_name)

        messagebox.showinfo("Report saved", f"Report saved as {out_name}")

    except Exception as e:
        messagebox.showerror("Error generating report", str(e))

# Add Generate Report button to Recommendations tab
generate_btn = tk.Button(tab_practitioner, text="Generate Report", command=generate_report, bg="#4CAF50", fg="white")
generate_btn.pack(anchor="e", padx=20, pady=20)


# ---------- End: Generate Report Functionality ----------






bind_mousewheel_to_patient()
root.mainloop()
