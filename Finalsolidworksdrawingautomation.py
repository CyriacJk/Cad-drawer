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
print("🚀 Launching SolidWorks...")
swApp = Dispatch("SldWorks.Application")
swApp.Visible = True

# === Open Assembly ===
print(f"📂 Opening assembly: {assembly_path}")
modelDoc = swApp.OpenDoc6(assembly_path, 2, 0, "", errors, warnings)
if not modelDoc:
    print(f"❌ Failed to open assembly. Error code: {errors.value}")
    exit()
print("✅ Assembly loaded.")

# === Gather part paths ===
configMgr = modelDoc.ConfigurationManager
active_config = configMgr.ActiveConfiguration
rootComponent = active_config.GetRootComponent3(True)

if not rootComponent:
    print("❌ Could not access root component.")
    exit()

children = rootComponent.GetChildren
part_paths = set()

for comp in children:
    model = comp.GetModelDoc2
    if model:
        path = model.GetPathName
        if path and path.lower().endswith(".sldprt"):
            part_paths.add(path)

print(f"\n🔍 Found {len(part_paths)} unique part(s).")

# === Helper: Try inserting any usable view ===
def try_insert_view(drawing_doc, file_path):
    possible_views = ["Default@Front", "Default@Isometric", "*Isometric", "*Front"]
    for view_key in possible_views:
        view = drawing_doc.CreateDrawViewFromModelView3(file_path, view_key, 0.2, 0.15, 0)
        if view:
            return True
    return False

# === Draw all parts ===
print("\n📐 Generating part drawings...")
for part_path in part_paths:
    part_name = os.path.splitext(os.path.basename(part_path))[0]
    drawing_path = os.path.join(output_dir, f"{part_name}.SLDDRW")

    part_doc = swApp.OpenDoc6(part_path, 1, 0, "", errors, warnings)
    if not part_doc:
        print(f"⚠️ Could not open part: {part_path}")
        continue

    drawing = swApp.NewDocument(drawing_template_path, 0, 0, 0)
    if drawing:
        time.sleep(1)
        if try_insert_view(drawing, part_path):
            drawing.Create3rdAngleViews(part_path)
            drawing.SaveAs(drawing_path)
            print(f"✅ Drawing saved: {drawing_path}")
        else:
            print(f"❌ Could not insert any view for: {part_path}")
    else:
        print(f"❌ Failed to create drawing for: {part_name}")

# === Draw full assembly ===
print("\n📐 Generating assembly drawing...")
assembly_name = os.path.splitext(os.path.basename(assembly_path))[0]
assembly_drawing_path = os.path.join(output_dir, f"{assembly_name}_assembly.SLDDRW")

drawing = swApp.NewDocument(drawing_template_path, 0, 0, 0)
if drawing:
    time.sleep(1)
    if try_insert_view(drawing, assembly_path):
        drawing.Create3rdAngleViews(assembly_path)
        drawing.SaveAs(assembly_drawing_path)
        print(f"✅ Assembly drawing saved: {assembly_drawing_path}")
    else:
        print("❌ Could not insert any view for the assembly.")
else:
    print("❌ Failed to create drawing for assembly.")

print("\n🎉 All drawings generated successfully!")
