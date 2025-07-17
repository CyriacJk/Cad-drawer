import os
import time
from win32com.client import Dispatch, VARIANT
import pythoncom

# === Paths ===
assembly_path = r"D:\PERSONAL\solidworks\solidworks cadd centre\Food chain assembly and motion study\table conveyor.SLDASM"
drawing_template_path = r"D:\PERSONAL\solidworks\template.drwdot"
output_dir = r"D:\PERSONAL\solidworks\solidworks cadd centre\Food chain assembly and motion study\Drawings"

# === Ensure output directory exists ===
os.makedirs(output_dir, exist_ok=True)

# === COM error containers ===
errors = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
warnings = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

# === Launch SolidWorks ===
print("\U0001F680 Launching SolidWorks...")
swApp = Dispatch("SldWorks.Application")
swApp.Visible = True

# === Open Assembly ===
print(f"\U0001F4C2 Opening assembly: {assembly_path}")
modelDoc = swApp.OpenDoc6(assembly_path, 2, 0, "", errors, warnings)
if not modelDoc:
    print(f"\u274C Failed to open assembly. Error code: {errors.value}")
    exit()
print("\u2705 Assembly loaded.")

# === Get component paths ===
configMgr = modelDoc.ConfigurationManager
active_config = configMgr.ActiveConfiguration
rootComponent = active_config.GetRootComponent3(True)

if not rootComponent:
    print("\u274C Could not access root component.")
    exit()

children = rootComponent.GetChildren
part_paths = set()

for comp in children:
    model = comp.GetModelDoc2
    if model:
        path = model.GetPathName
        if path and path.lower().endswith(".sldprt"):
            part_paths.add(path)

print(f"\n\U0001F50D Found {len(part_paths)} unique part(s).")

# === Function to run DimXpert ===
def autodim_part(part_doc):
    try:
        ext = part_doc.Extension()
        dx_mgr = getattr(ext, "MBD Dimensions", None)
        if dx_mgr is None:
            print("⚠️  DimXpertManager not available for this part or license.")
            return
        dx_mgr.AutoDimensionScheme("ANSI", 0, 0)  # Change "ANSI" to "ISO" if needed
        part_doc.EditRebuild3()   
        print("✅ DimXpert applied successfully.")
    except Exception as e:
        print(f"❌ DimXpert failed: {e}")

# === Function to insert DimXpert annotations to drawing ===
def import_dimxpert_to_drawing(drawing_doc):
    try:
        swInsertDimXpert = 512   # Flag for DimXpert annotations
        swAllViews = 4095        # All views
        drawing_doc.InsertModelItems(swInsertDimXpert, swAllViews)
        print("✅ Inserted DimXpert annotations to drawing.")
    except Exception as e:
        print(f"❌ Failed to import DimXpert annotations: {e}")

# === Generate drawings for each part ===
print("\n\U0001F4C0 Generating part drawings with GD&T...")
for part_path in part_paths:
    part_name = os.path.splitext(os.path.basename(part_path))[0]
    drawing_path = os.path.join(output_dir, f"{part_name}.SLDDRW")

    part_doc = swApp.OpenDoc6(part_path, 1, 0, "", errors, warnings)
    if not part_doc:
        print(f"⚠️  Could not open part: {part_path}")
        continue

    autodim_part(part_doc)

    drawing = swApp.NewDocument(drawing_template_path, 0, 0, 0)
    if drawing:
        time.sleep(1)
        drawing.Create3rdAngleViews(part_path)
        import_dimxpert_to_drawing(drawing)
        drawing.SaveAs(drawing_path)
        print(f"✅ GD&T drawing saved: {drawing_path}")
    else:
        print(f"❌ Failed to create drawing for: {part_name}")

# === Assembly drawing (no GD&T unless manually applied) ===
print("\n\U0001F4C0 Generating assembly drawing...")
assembly_name = os.path.splitext(os.path.basename(assembly_path))[0]
assembly_drawing_path = os.path.join(output_dir, f"{assembly_name}_assembly.SLDDRW")

drawing = swApp.NewDocument(drawing_template_path, 0, 0, 0)
if drawing:
    time.sleep(1)
    drawing.Create3rdAngleViews(assembly_path)
    drawing.SaveAs(assembly_drawing_path)
    print(f"✅ Assembly drawing saved: {assembly_drawing_path}")
else:
    print("❌ Failed to create drawing for assembly.")

print("\n\U0001F389 All drawings generated successfully with GD&T!")
