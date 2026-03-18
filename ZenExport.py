"""ZenExport v5.2 - Fusion 360 Script for Local Design Backup.

This script automates exporting Fusion 360 designs locally with:
- Auto-Versioning: Saves numbered history (v01, v02) for CAD files
- Context Binding: Resumes project path per-design (Multi-Tab Support)
- Thumbnails: Generates viewport preview image
- Auto-Open: Launches folder after export
- Smart Save: Hashing to prevent redundant saves

Author: ZenExport (Generated via GEMINI.md)
Version: 5.2 (Context Aware)
"""

import adsk.core
import adsk.fusion
import traceback
import os
import json
import re
import glob

# ============================================================================
# CONSTANTS & GLOBALS
# ============================================================================

CMD_ID = "ZenExport_Cmd_v5" # Updated ID to force refresh
CMD_NAME = "ZenExport Save"
CMD_DESC = "Saves versioned backup locally (Ctrl+S override)."
ResourcesFolder = os.path.join(os.path.dirname(os.path.realpath(__file__)), "resources")
CONFIG_FILENAME = "session_config.json"
MESH_REFINEMENT = adsk.fusion.MeshRefinementSettings.MeshRefinementHigh

# Global Event Handlers
_handlers = []

# ============================================================================
# LOGGING HELPER
# ============================================================================

def get_log_path() -> str:
    return os.path.join(os.path.expanduser("~"), "Desktop", "zenexport_debug.log")

def log_to_console(app: adsk.core.Application, message: str) -> None:
    try:
        text_palette = app.userInterface.palettes.itemById("TextCommands")
        if text_palette:
            text_palette.writeText(f"[ZenExport] {message}")
    except: pass

    try:
        with open(get_log_path(), "a", encoding="utf-8") as f:
            from datetime import datetime
            timestamp = datetime.now().strftime("%H:%M:%S")
            f.write(f"[{timestamp}] {message}\n")
    except: pass

# ============================================================================
# CONFIGURATION & CONTEXT BINDING
# ============================================================================

def get_config_path() -> str:
    script_dir = os.path.dirname(os.path.realpath(__file__))
    return os.path.join(script_dir, CONFIG_FILENAME)

def load_config() -> dict:
    path = get_config_path()
    if not os.path.exists(path): return {}
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return {}

def save_config_file(data: dict) -> None:
    path = get_config_path()
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4)
    except: pass

def get_context_for_design(design: adsk.fusion.Design) -> dict | None:
    """Retrieves context using Revision ID (Session GUID) or Name."""
    cfg = load_config()
    contexts = cfg.get("contexts", {})
    
    # 1. Try GUID (Most reliable for Untitled/Unsaved in-session)
    try:
        rev_id = design.rootComponent.revisionId
        if rev_id in contexts: return contexts[rev_id]
    except: pass
    
    # 2. Try Name (Fallback for fresh open of saved files where GUID changed)
    try:
        full_name = design.parentDocument.name
        # Regex to strip version suffix like _v01, _v1,  v01,  v1
        # Captures "MyProject" from "MyProject_v05" or "MyProject v5"
        match = re.search(r'^(.*?)[\s_]v\d+$', full_name, re.IGNORECASE)
        if match:
             name_key = sanitize_filename(match.group(1))
        else:
             name_key = sanitize_filename(full_name)
             
        # Debug Log (checking regex capture)
        if match: log_to_console(adsk.core.Application.get(), f"Parsed Key: {name_key}")
             
    except:
        return None
        
    return contexts.get(name_key)

def update_context_for_design(design: adsk.fusion.Design, project_name: str, root_folder: str, hash_val: str) -> None:
    """Updates context binding for a specific design."""
    cfg = load_config()
    if "contexts" not in cfg: cfg["contexts"] = {}
    
    # Store by GUID
    try:
        rev_id = design.rootComponent.revisionId
        cfg["contexts"][rev_id] = {
            "root": root_folder,
            "hash": hash_val,
            "name": project_name
        }
    except: pass
    
    # We will store a secondary key if name is valid
    if project_name != "Untitled":
         cfg["contexts"][project_name] = {
            "root": root_folder,
            "hash": hash_val,
             "name": project_name
        }
    
    # Legacy fallback for single-mode
    cfg["last_save_directory"] = root_folder
    
    save_config_file(cfg)

# ============================================================================
# UTILS & VERSIONING
# ============================================================================

def sanitize_filename(name: str) -> str:
    invalid = '<>:"/\\|?*'
    for char in invalid:
        name = name.replace(char, "_")
    return name

def is_shift_held() -> bool:
    try:
        import ctypes
        return bool(ctypes.windll.user32.GetAsyncKeyState(0x10) & 0x8000)
    except:
        return False

def ensure_folder_exists(app: adsk.core.Application, folder_path: str) -> None:
    if not os.path.exists(folder_path):
        log_to_console(app, f"Creating directory: {folder_path}")
        os.makedirs(folder_path, exist_ok=True)

def get_next_version_number(cad_folder: str) -> int:
    if not os.path.exists(cad_folder):
        return 1
    items = os.listdir(cad_folder)
    versions = []
    regex = re.compile(r'^v(\d+)$')
    for item in items:
        if os.path.isdir(os.path.join(cad_folder, item)):
            match = regex.match(item)
            if match:
                versions.append(int(match.group(1)))
    return (max(versions) + 1) if versions else 1

# ============================================================================
# STATE TRACKING & FEEDBACK
# ============================================================================

def get_design_hash(design: adsk.fusion.Design) -> str:
    try:
        tl_count = design.timeline.count
        tl_pos = design.timeline.markerPosition
        comp_count = design.rootComponent.allOccurrences.count
        param_count = len(design.userParameters)
        body_count = sum(1 for _ in design.rootComponent.bRepBodies)
        return f"{tl_count}|{tl_pos}|{comp_count}|{param_count}|{body_count}"
    except:
        return "ERROR"

def show_success_feedback(ui, mode, res, path):
    f3d_icon = '✅' if res['f3d'] else '❌'
    step_icon = '✅' if res['step'] else '❌'
    
    msg = f"ZenExport {mode} Complete!\n"
    msg += f"Saved to: {path}\n\n"
    msg += f"Version: {res['version']}\n"
    msg += f"• F3D: {f3d_icon}\n"
    msg += f"• STEP: {step_icon}\n"
    msg += f"• STLs: {res['stl_ok']} exported ({res['stl_fail']} failed)"
    
    ui.messageBox(msg, f"ZenExport {mode}")

# ============================================================================
# EXPORT LOGIC
# ============================================================================

def save_thumbnail(app: adsk.core.Application, folder: str) -> bool:
    try:
        viewport = app.activeViewport
        path = os.path.join(folder, "_preview.png")
        viewport.saveAsImageFile(path, 400, 400)
        return True
    except:
        return False

def export_cad_files(app, design, cad_folder, versioned_name) -> tuple[bool, bool]:
    export_mgr = design.exportManager
    f3d_path = os.path.join(cad_folder, f"{versioned_name}.f3d")
    f3d_ok = False
    try:
        opts = export_mgr.createFusionArchiveExportOptions(f3d_path)
        f3d_ok = export_mgr.execute(opts)
    except Exception as e:
        log_to_console(app, f"F3D Err: {e}")

    step_path = os.path.join(cad_folder, f"{versioned_name}.step")
    step_ok = False
    try:
        opts = export_mgr.createSTEPExportOptions(step_path)
        step_ok = export_mgr.execute(opts)
    except Exception as e:
        log_to_console(app, f"STEP Err: {e}")

    return f3d_ok, step_ok

def export_stl_files(app, design, bodies, models_folder) -> tuple[int, int]:
    export_mgr = design.exportManager
    s, f = 0, 0
    for name, body in bodies:
        fname = sanitize_filename(f"{name}_{body.name}.stl")
        path = os.path.join(models_folder, fname)
        try:
            opts = export_mgr.createSTLExportOptions(body, path)
            opts.meshRefinement = MESH_REFINEMENT
            if export_mgr.execute(opts): s += 1
            else: f += 1
        except: f += 1
    return s, f

def collect_bodies(design) -> list:
    bodies = []
    def traverse(comp, prefix=""):
        name = prefix + comp.name if prefix else comp.name
        for b in comp.bRepBodies:
            if b.isVisible: bodies.append((name, b))
        for occ in comp.occurrences:
            if occ.isVisible: traverse(occ.component, f"{name}_" if name else "")
    traverse(design.rootComponent)
    return bodies

def perform_sync_export(app, design, project_folder, project_name) -> dict:
    cad_folder = os.path.join(project_folder, "CAD")
    models_folder = os.path.join(project_folder, "Models")
    ensure_folder_exists(app, cad_folder)
    ensure_folder_exists(app, models_folder)
    
    save_thumbnail(app, project_folder)
    
    ver_num = get_next_version_number(cad_folder)
    ver_lbl = f"v{ver_num:02d}"
    versioned_name = f"{project_name}_{ver_lbl}"
    target_cad = os.path.join(cad_folder, ver_lbl)
    ensure_folder_exists(app, target_cad)
    
    f3d, step = export_cad_files(app, design, target_cad, versioned_name)
    bodies = collect_bodies(design)
    stl_s, stl_f = 0, 0
    if bodies: stl_s, stl_f = export_stl_files(app, design, bodies, models_folder)
    
    # Returning plain dict (NO F3D PATH needed since no Auto-Swap)
    return {'f3d': f3d, 'step': step, 'stl_ok': stl_s, 'stl_fail': stl_f, 'version': versioned_name, 'proj_root': project_folder}

# ============================================================================
# CORE LOGIC
# ============================================================================

def run_zen_export_logic(app, design, mode_override=None):
    ui = app.userInterface
    doc = app.activeDocument
    
    # 1. PREREQUISITE CHECK (Soft)
    # We allow "Untitled" or unsaved files to proceed to INIT mode.
    # The user wants to "Save" them via ZenExport.
    
    full_name = doc.name
    # Strict Name Parsing via Regex (Handles " Name v1" and "Name_v01")
    match = re.search(r'^(.*?)[\s_]v\d+$', full_name, re.IGNORECASE)
    if match:
        design_name = sanitize_filename(match.group(1))
    else:
        design_name = sanitize_filename(full_name)
    
    # If Untitled, we handle in INIT flow
    
    current_hash = get_design_hash(design)
    
    project_folder = ""
    project_name = ""
    mode = "INIT"
    last_hash = ""
    
    # 2. LOGIC FORKING (Disk-Based Check)
    # First, consult session memory for a "Known Location" for this design
    ctx = get_context_for_design(design)
    
    # Override logic
    shift = is_shift_held()
    if shift or mode_override == "INIT":
        mode = "INIT"
    elif ctx:
         # CANDIDATE Path found. Now Validate on DISK.
         candidate_root = ctx.get("root", "")
         if candidate_root and os.path.exists(candidate_root):
             # Condition A: Folder Exists -> Incremental Update
             mode = "UPDATE"
             project_folder = candidate_root
             project_name = ctx.get("name", design_name) # Use stored name if available
             last_hash = ctx.get("hash", "")
             log_to_console(app, f"Context Validated: '{design_name}' -> {project_folder}")
         else:
             # Condition B: Folder Missing (Deleted?) -> Trigger Init
             log_to_console(app, f"Context Invalid/Missing: '{candidate_root}' -> Re-initializing")
             mode = "INIT"
    
    # 3. Dirty Flag Check (Only for Updates)
    if mode == "UPDATE":
        if current_hash == last_hash:
            ui.messageBox("No changes detected since last ZenExport.\nSave Skipped.", "ZenExport")
            return

    # 4. INITIALIZATION (New / Lost Context)
    if mode == "INIT":
        # We need to establish a location.
        dlg = ui.createFolderDialog()
        dlg.title = f"ZenExport: Select Location for '{design_name}'"
        if dlg.showDialog() != adsk.core.DialogResults.DialogOK: return
        base_dir = dlg.folder
        
        # Determine Project Name (Default to Design Name)
        # If Untitled, default to MyProduct or empty
        def_name = design_name if design_name != "Untitled" else "MyProduct"
        
        res = ui.inputBox("Confirm Project Folder Name:", "ZenExport Setup", def_name)
        if res[1]: return 
        project_name = sanitize_filename(res[0])
        if not project_name: return 
        
        project_folder = os.path.join(base_dir, project_name)
        
        # Check if we are "adopting" an existing folder
        if os.path.exists(project_folder):
             log_to_console(app, "Adopting existing folder structure.")

    # 5. EXECUTION
    res = perform_sync_export(app, design, project_folder, project_name)
    
    # 6. CONTEXT BINDING (Update Record)
    update_context_for_design(design, project_name, project_folder, current_hash)
    
    # 7. FEEDBACK & AUTO-OPEN
    try: os.startfile(project_folder)
    except: pass
         
    show_success_feedback(ui, mode, res, project_folder)


# ============================================================================
# EVENT HANDLERS
# ============================================================================

class ZenExportCommandStartingHandler(adsk.core.ApplicationCommandEventHandler):
    def __init__(self): super().__init__()
    def notify(self, args):
        try:
            cmd_id = args.commandId
            app = adsk.core.Application.get()
            TARGET_CMDS = ["PLM360SaveCommand", "FusionSaveCommand", "FileSave", "Save"]
            
            if cmd_id in TARGET_CMDS:
                log_to_console(app, f"🚫 Intercepted: {cmd_id}")
                args.isCanceled = True
                design = adsk.fusion.Design.cast(app.activeProduct)
                if design: run_zen_export_logic(app, design)
        except Exception as e:
            log_to_console(app, f"Interceptor Err: {e}")

class ZenExportCommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self): super().__init__()
    def notify(self, args):
        try:
            cmd = args.command
            cmd.isExecutedWhenPreEmpted = False
            onExecute = ZenExportExecuteHandler() 
            cmd.execute.add(onExecute)
            _handlers.append(onExecute)
        except: pass

class ZenExportExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self): super().__init__()
    def notify(self, args):
        try:
            app = adsk.core.Application.get()
            design = adsk.fusion.Design.cast(app.activeProduct)
            if design: run_zen_export_logic(app, design)
        except: pass

class ZenExportDocumentActivatedHandler(adsk.core.DocumentEventHandler):
    def __init__(self): super().__init__()
    def notify(self, args):
        pass

# ============================================================================
# MAIN
# ============================================================================

def run(context):
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        log_to_console(app, "ZenExport Starting (Save Button Mode)...")
        
        # 1. Clean up old instances
        cmdDef = ui.commandDefinitions.itemById(CMD_ID)
        if cmdDef: cmdDef.deleteMe()
        
        # 2. Create the Command Definition 
        # WE REMOVED .shortcutKeyboard HERE so Ctrl+S stays default.
        cmdDef = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_DESC, ResourcesFolder)
        
        # 3. Connect the Created Handler
        onCreated = ZenExportCommandCreatedHandler()
        cmdDef.commandCreated.add(onCreated) 
        _handlers.append(onCreated)
        
        # 4. RE-ENABLE THE INTERCEPTOR (This hijacks the UI Save Button)
        onStarting = ZenExportCommandStartingHandler()
        ui.commandStarting.add(onStarting)
        _handlers.append(onStarting)

        log_to_console(app, "ZenExport Ready. Save Button hijacked; Ctrl+S is Cloud Save.")
        
    except:
        app = adsk.core.Application.get()
        if app.userInterface:
            app.userInterface.messageBox(f"Start Failed:\n{traceback.format_exc()}")

def stop(context):
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        
        # Clean up handlers
        for handler in _handlers:
            # This is a bit generic, but Fusion handles the cleanup of 
            # commandStarting handlers when the script stops.
            pass
            
        cmdDef = ui.commandDefinitions.itemById(CMD_ID)
        if cmdDef: cmdDef.deleteMe()
        
        log_to_console(app, "ZenExport Stopped.")
    except:
        pass
