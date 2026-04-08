import adsk.core, adsk.fusion, traceback, os, pathlib, webbrowser

# Global list to keep event handlers active
handlers = []

def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        design = adsk.fusion.Design.cast(app.activeProduct)
        if not design:
            ui.messageBox('No active design found.')
            return

        # Command ID updated to ModingeFusion360Exporter
        cmdDef = ui.commandDefinitions.itemById('ModingeFusion360Exporter')
        if not cmdDef:
            cmdDef = ui.commandDefinitions.addButtonDefinition('ModingeFusion360Exporter', 'Export Files', '')

        # --- EVENT 1: MAIN CONFIGURATION WINDOW ---
        class DialogCreatedHandler(adsk.core.CommandCreatedEventHandler):
            def __init__(self):
                super().__init__()
            def notify(self, args):
                try:
                    cmd = args.command
                    inputs = cmd.commandInputs

                    # MODINGE Logo
                    script_dir = os.path.dirname(os.path.realpath(__file__))
                    logo_path = os.path.join(script_dir, 'recursos', 'MODINGE-logo-19.png')
                    logo_uri = pathlib.Path(logo_path).as_uri()

                    html_logo = f'<div align="center"><img src="{logo_uri}" height="85"></div>'
                    txt_logo = inputs.addTextBoxCommandInput('img_banner', '', html_logo, 4, True)
                    txt_logo.isFullWidth = True

                    # 1. Basic Settings
                    txt_name = inputs.addTextBoxCommandInput('txt_step1', '', '<b>1. Basic Settings</b>', 1, True)
                    txt_name.isFullWidth = True
                    
                    default_name = design.parentDocument.name if design.parentDocument else 'Export'
                    inputs.addStringValueInput('base_name', 'Base name:', default_name)

                    btn_browse = inputs.addBoolValueInput('btn_browse', 'Location:', False, '', False)
                    btn_browse.text = '📂 Browse Folder...'
                    
                    txt_path = inputs.addStringValueInput('txt_path', 'Destination:', 'Select a folder...')
                    txt_path.isReadOnly = True

                    # 2. Formats
                    format_group = inputs.addGroupCommandInput('format_group', '2. Export Formats')
                    format_group.isExpanded = True 
                    format_group.children.addBoolValueInput('chk_stl', 'Export .stl', True, '', True)
                    format_group.children.addBoolValueInput('chk_3mf', 'Export .3mf', True, '', True)
                    format_group.children.addBoolValueInput('chk_f3d', 'Export .f3d', True, '', True)
                    format_group.children.addBoolValueInput('chk_step', 'Export .step', True, '', True)
                    format_group.children.addBoolValueInput('chk_iges', 'Export .iges', True, '', True)

                    # 3. Scope
                    scope_group = inputs.addGroupCommandInput('scope_group', '3. What do you want to export?')
                    scope_group.isExpanded = True
                    scope_group.children.addBoolValueInput('chk_scope_root', 'Entire Design (Root)', True, '', True)
                    scope_group.children.addBoolValueInput('chk_scope_bodies', 'Individual Bodies', True, '', True)

                    # Note
                    html_footer = '<div style="color:#d9534f; font-size: 11px;"><i><b>Note:</b> Exporting many bodies in multiple formats may take some time.</i></div><hr>'
                    txt_footer = inputs.addTextBoxCommandInput('txt_footer', '', html_footer, 1, True)
                    txt_footer.isFullWidth = True

                    # Main Contact Button
                    btn_bio = inputs.addBoolValueInput('btn_link_bio', '', False, '', False)
                    btn_bio.text = '📱 Contact Information'
                    btn_bio.isFullWidth = True

                    # Handlers
                    onInputChanged = UIInputChangedHandler()
                    cmd.inputChanged.add(onInputChanged)
                    handlers.append(onInputChanged)

                    onExecute = ExecuteHandler()
                    cmd.execute.add(onExecute)
                    handlers.append(onExecute)

                except:
                    if ui: ui.messageBox('Error creating interface:\n{}'.format(traceback.format_exc()))

        class UIInputChangedHandler(adsk.core.InputChangedEventHandler):
            def __init__(self):
                super().__init__()
            def notify(self, args):
                try:
                    cmdInput = args.input
                    if cmdInput.id == 'btn_browse':
                        folderDlg = ui.createFolderDialog()
                        folderDlg.title = 'Select destination folder'
                        if folderDlg.showDialog() == adsk.core.DialogResults.DialogOK:
                            args.inputs.itemById('txt_path').value = folderDlg.folder
                    elif cmdInput.id == 'btn_link_bio':
                        webbrowser.open('https://lnk.bio/ModInge')
                except:
                    pass

        class ExecuteHandler(adsk.core.CommandEventHandler):
            def __init__(self):
                super().__init__()
            def notify(self, args):
                try:
                    inputs = args.command.commandInputs
                    base_name = inputs.itemById('base_name').value
                    dest_path = inputs.itemById('txt_path').value

                    if dest_path == 'Select a folder...' or not os.path.exists(dest_path):
                        ui.messageBox('Please select a valid destination folder using the "Browse Folder..." button.')
                        args.executeFailed = True
                        return

                    formats_to_process = []
                    if inputs.itemById('chk_stl').value: formats_to_process.append('.stl')
                    if inputs.itemById('chk_3mf').value: formats_to_process.append('.3mf')
                    if inputs.itemById('chk_f3d').value: formats_to_process.append('.f3d')
                    if inputs.itemById('chk_step').value: formats_to_process.append('.step')
                    if inputs.itemById('chk_iges').value: formats_to_process.append('.iges')

                    exp_root = inputs.itemById('chk_scope_root').value
                    exp_bodies = inputs.itemById('chk_scope_bodies').value

                    if not formats_to_process:
                        ui.messageBox("You must select at least one format.")
                        args.executeFailed = True
                        return
                    if not exp_root and not exp_bodies:
                        ui.messageBox("You must select at least one export scope (Root or Bodies).")
                        args.executeFailed = True
                        return

                    exportMgr = design.exportManager

                    def run_export(entity, final_name, is_body):
                        for ext in formats_to_process:
                            folder_name = f"{ext.upper()} - {base_name}"
                            path_dir = os.path.join(dest_path, folder_name)
                            if not os.path.exists(path_dir): os.makedirs(path_dir)
                            full_path = os.path.join(path_dir, f"{final_name}{ext}")
                            try:
                                if ext == '.stl': opts = exportMgr.createSTLExportOptions(entity, full_path)
                                elif ext == '.3mf': opts = exportMgr.createC3MFExportOptions(entity, full_path)
                                elif ext == '.f3d':
                                    if is_body: continue
                                    opts = exportMgr.createFusionArchiveExportOptions(full_path)
                                elif ext == '.step': opts = exportMgr.createSTEPExportOptions(full_path)
                                elif ext == '.iges': opts = exportMgr.createIGESExportOptions(full_path)
                                exportMgr.execute(opts)
                                adsk.doEvents()
                            except: continue

                    if exp_root: run_export(design.rootComponent, base_name, False)
                    if exp_bodies:
                        for comp in design.allComponents:
                            occurrences = design.rootComponent.allOccurrencesByComponent(comp)
                            if occurrences.count > 0:
                                assembly_visible = False
                                for occ in occurrences:
                                    if occ.isVisible:
                                        assembly_visible = True
                                        break
                                if not assembly_visible:
                                    continue 

                            for body in comp.bRepBodies:
                                if body.isVisible:
                                    clean_name = "".join([c for c in f"{comp.name}_{body.name}" if c.isalnum() or c in (' ', '_', '-')])
                                    run_export(body, clean_name, True)

                    # When finished, trigger the custom success window
                    show_success_dialog(ui)

                except:
                    if ui: ui.messageBox('Error exporting:\n{}'.format(traceback.format_exc()))

        # Start main command
        onCommandCreated = DialogCreatedHandler()
        cmdDef.commandCreated.add(onCommandCreated)
        handlers.append(onCommandCreated)
        cmdDef.execute()
        adsk.autoTerminate(False)

    except:
        if ui: ui.messageBox('Error:\n{}'.format(traceback.format_exc()))

# --- NEW FUNCTION: FINAL SUCCESS WINDOW ---
def show_success_dialog(ui):
    try:
        cmdDefFinal = ui.commandDefinitions.itemById('ModingeFusion360Exporter_Final')
        if not cmdDefFinal:
            cmdDefFinal = ui.commandDefinitions.addButtonDefinition('ModingeFusion360Exporter_Final', 'Process Completed', '')

        class FinalDialogCreatedHandler(adsk.core.CommandCreatedEventHandler):
            def __init__(self):
                super().__init__()
            def notify(self, args):
                cmd = args.command
                cmd.okButtonText = 'Close'
                inputs = cmd.commandInputs

                # Success message
                html_msg = (
                    '<div align="center">'
                    '<p style="font-size: 13px;"><b>Export Successful!</b></p>'
                    '<p>The files have been correctly organized.</p>'
                    '<p style="font-size: 11px;">If you have any questions or need technical support, feel free to contact me.</p>'
                    '</div>'
                )
                txt = inputs.addTextBoxCommandInput('txt_final', '', html_msg, 4, True)
                txt.isFullWidth = True

                # Contact button again
                btn_final = inputs.addBoolValueInput('btn_link_final', '', False, '', False)
                btn_final.text = '📱 Contact Information'
                btn_final.isFullWidth = True

                # Connect button click
                onInputChangedFinal = UIInputChangedFinalHandler()
                cmd.inputChanged.add(onInputChangedFinal)
                handlers.append(onInputChangedFinal)

        class UIInputChangedFinalHandler(adsk.core.InputChangedEventHandler):
            def __init__(self):
                super().__init__()
            def notify(self, args):
                if args.input.id == 'btn_link_final':
                    webbrowser.open('https://lnk.bio/ModInge')

        onCreatedFinal = FinalDialogCreatedHandler()
        cmdDefFinal.commandCreated.add(onCreatedFinal)
        handlers.append(onCreatedFinal)
        cmdDefFinal.execute()

    except:
        pass