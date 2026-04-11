import adsk.core, adsk.fusion, traceback
import pathlib, webbrowser, csv, os

handlers = []
app = adsk.core.Application.get()
ui = app.userInterface
success_flag = False

# Custom Event ID
CUSTOM_EVENT_ID = 'ModingeFusion360_Success_Trigger_Event'

# ==========================================
# FILTERING & PROCESSING LOGIC
# ==========================================

def get_filtered_parameters(only_favorites):
    design = adsk.fusion.Design.cast(app.activeProduct)
    if not design: return []
    
    if only_favorites:
        return [p for p in design.allParameters if p.isFavorite]
    else:
        return [p for p in design.allParameters]

def export_csv(only_favorites, delim_char):
    params = get_filtered_parameters(only_favorites)
    if not params:
        ui.messageBox('No parameters found to export.')
        return False
    
    fileDialog = ui.createFileDialog()
    fileDialog.title = "Export MODINGE Selection"
    fileDialog.filter = "CSV Files (*.csv)"
    
    if fileDialog.showSave() == adsk.core.DialogResults.DialogOK:
        filename = fileDialog.filename
        # utf-8-sig adds the BOM marker that Excel needs to recognize special characters properly
        with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=delim_char) 
            writer.writerow(['Name', 'Unit', 'Expression', 'Comment', 'Favorite'])
            for p in params:
                writer.writerow([p.name, p.unit, p.expression, p.comment, str(p.isFavorite).upper()])
        return True
    return False

def import_csv():
    design = adsk.fusion.Design.cast(app.activeProduct)
    if not design: return False
    
    fileDialog = ui.createFileDialog()
    fileDialog.title = "Import MODINGE Parameters"
    fileDialog.filter = "CSV Files (*.csv)"
    
    if fileDialog.showOpen() == adsk.core.DialogResults.DialogOK:
        filename = fileDialog.filename
        # utf-8-sig safely strips hidden Excel BOM characters on import
        with open(filename, 'r', encoding='utf-8-sig') as f:
            sample = f.read(2048)
            f.seek(0)
            
            # Robust auto-detection of the delimiter
            try:
                delim = csv.Sniffer().sniff(sample, delimiters=',;\t|').delimiter
            except:
                first_line = sample.split('\n')[0]
                delim = ';' if ';' in first_line else ','
            
            reader = csv.reader(f, delimiter=delim)
            next(reader) # Skip header
            
            for row in reader:
                if len(row) >= 3:
                    name, unit, expr = row[0].strip(), row[1].strip(), row[2].strip()
                    comment = row[3].strip() if len(row) > 3 else ''
                    
                    is_fav = False
                    if len(row) > 4:
                        fav_str = row[4].strip().upper()
                        # Multi-language support for the boolean flag
                        is_fav = fav_str in ['TRUE', 'VERDADERO', 'T', 'V', '1']
                    
                    param = design.allParameters.itemByName(name)
                    if param:
                        try:
                            param.expression = expr
                            param.comment = comment
                            param.isFavorite = is_fav
                        except: pass
                    else:
                        try:
                            new_p = design.userParameters.add(name, adsk.core.ValueInput.createByString(expr), unit, comment)
                            new_p.isFavorite = is_fav
                        except: pass
        return True
    return False

# ==========================================
# EXECUTION EVENTS & CUSTOM BRIDGE
# ==========================================

class SuccessTriggerHandler(adsk.core.CustomEventHandler):
    def notify(self, args):
        try: launch_success_dialog()
        except: pass

class MainCommandExecuteHandler(adsk.core.CommandEventHandler):
    def notify(self, args):
        global success_flag
        try:
            inputs = args.command.commandInputs
            is_export = inputs.itemById('chk_export').value
            only_favs = inputs.itemById('chk_favs').value
            
            # Read the delimiter selected by the user
            delim_selection = inputs.itemById('delim_format').selectedItem.name
            delim_char = ';' if 'Semicolon' in delim_selection else ','
            
            if is_export:
                if export_csv(only_favs, delim_char): success_flag = True
            else:
                if import_csv(): success_flag = True
        except:
            ui.messageBox('Execution Error:\n{}'.format(traceback.format_exc()))

class MainCommandDestroyHandler(adsk.core.CommandEventHandler):
    def notify(self, args):
        global success_flag
        try:
            if success_flag:
                success_flag = False
                app.fireCustomEvent(CUSTOM_EVENT_ID)
            else:
                adsk.terminate()
        except: pass

class MainInputChangedHandler(adsk.core.InputChangedEventHandler):
    def notify(self, args):
        try:
            eventArgs = adsk.core.InputChangedEventArgs.cast(args)
            id = eventArgs.input.id
            inputs = eventArgs.inputs

            if id == 'contact_btn':
                webbrowser.open('https://lnk.bio/ModInge')
            
            elif id == 'chk_export' and eventArgs.input.value:
                inputs.itemById('chk_import').value = False
            elif id == 'chk_import' and eventArgs.input.value:
                inputs.itemById('chk_export').value = False
            
            elif id == 'chk_favs' and eventArgs.input.value:
                inputs.itemById('chk_all').value = False
            elif id == 'chk_all' and eventArgs.input.value:
                inputs.itemById('chk_favs').value = False

            if id in ['chk_export', 'chk_import']:
                if not inputs.itemById('chk_export').value and not inputs.itemById('chk_import').value:
                    eventArgs.input.value = True
            
            if id in ['chk_favs', 'chk_all']:
                if not inputs.itemById('chk_favs').value and not inputs.itemById('chk_all').value:
                    eventArgs.input.value = True

        except: pass

# ==========================================
# USER INTERFACE (MODINGE UI)
# ==========================================

class MainCommandCreatedEventHandler(adsk.core.CommandCreatedEventHandler):
    def notify(self, args):
        try:
            cmd = args.command
            cmd.okButtonText = 'Execute Process'
            inputs = cmd.commandInputs
            
            # BANNER
            logo_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'recursos', 'MODINGE-logo-19.png')
            logo_uri = pathlib.Path(logo_path).as_uri()
            inputs.addTextBoxCommandInput('banner', '', f'<img src="{logo_uri}" height="85">', 4, True).isFullWidth = True

            # SECTION 1: ACTION
            inputs.addTextBoxCommandInput('title1', '', '<b>Select Action:</b>', 1, True).isFullWidth = True
            inputs.addBoolValueInput('chk_export', '📤 Export to CSV', True, '', True)
            inputs.addBoolValueInput('chk_import', '📥 Import from CSV', True, '', False)

            # SECTION 2: SCOPE
            inputs.addTextBoxCommandInput('space1', '', '<br>', 1, True).isFullWidth = True
            inputs.addTextBoxCommandInput('title2', '', '<b>Document Scope:</b>', 1, True).isFullWidth = True
            inputs.addBoolValueInput('chk_favs', '★ Only Favorite Parameters', True, '', True)
            inputs.addBoolValueInput('chk_all', '📄 All Document Parameters', True, '', False)

            # SECTION 3: CSV FORMAT SETTINGS (New Feature)
            inputs.addTextBoxCommandInput('space_delim', '', '<br>', 1, True).isFullWidth = True
            inputs.addTextBoxCommandInput('title_delim', '', '<b>CSV Delimiter (Excel Compatibility):</b>', 1, True).isFullWidth = True
            
            delim_drop = inputs.addDropDownCommandInput('delim_format', '', adsk.core.DropDownStyles.TextListDropDownStyle)
            delim_drop.listItems.add('Comma (,) - Global Standard', True)
            delim_drop.listItems.add('Semicolon (;) - European / Spanish Excel', False)
            delim_drop.isFullWidth = True

            # FOOTER
            inputs.addTextBoxCommandInput('space2', '', '<br>', 1, True).isFullWidth = True
            inputs.addTextBoxCommandInput('footer', '', '⚠️ Note: Configure your process and click "Execute Process".', 1, True).isFullWidth = True

            # CONTACT BUTTON
            contact = inputs.addBoolValueInput('contact_btn', '', False, '', True)
            contact.text = '📱 Contact Information'
            contact.isFullWidth = True

            on_execute = MainCommandExecuteHandler()
            cmd.execute.add(on_execute)
            handlers.append(on_execute)

            on_destroy = MainCommandDestroyHandler()
            cmd.destroy.add(on_destroy)
            handlers.append(on_destroy)

            on_input = MainInputChangedHandler()
            cmd.inputChanged.add(on_input)
            handlers.append(on_input)

        except: ui.messageBox(traceback.format_exc())

# ==========================================
# SUCCESS DIALOG & CLEAN SHUTDOWN
# ==========================================

class SuccessCommandDestroyHandler(adsk.core.CommandEventHandler):
    def notify(self, args):
        try: adsk.terminate()
        except: pass

def launch_success_dialog():
    cmdDef = ui.commandDefinitions.itemById('ModingeFusion360_Success')
    if cmdDef: cmdDef.deleteMe()
    
    cmdDef = ui.commandDefinitions.addButtonDefinition('ModingeFusion360_Success', 'Success', '')
    on_created = SuccessCommandCreatedHandler()
    cmdDef.commandCreated.add(on_created)
    handlers.append(on_created)
    cmdDef.execute()

class SuccessCommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def notify(self, args):
        cmd = args.command
        cmd.isOKButtonVisible = False
        cmd.cancelButtonText = 'Close'
        
        txt = '<b>Synchronization Successful! 🎉</b><br>The model has been updated correctly.'
        args.command.commandInputs.addTextBoxCommandInput('s', '', txt, 3, True).isFullWidth = True
        
        btn = args.command.commandInputs.addBoolValueInput('contact_success', '', False, '', True)
        btn.text = '📱 Contact Information'
        btn.isFullWidth = True
        
        class SuccessContactHandler(adsk.core.InputChangedEventHandler):
            def notify(self, args):
                if args.input.id == 'contact_success': webbrowser.open('https://lnk.bio/ModInge')
        
        on_changed = SuccessContactHandler()
        cmd.inputChanged.add(on_changed)
        handlers.append(on_changed)

        on_destroy = SuccessCommandDestroyHandler()
        cmd.destroy.add(on_destroy)
        handlers.append(on_destroy)

# ==========================================
# SCRIPT INITIALIZATION
# ==========================================

def run(context):
    try:
        adsk.autoTerminate(False)

        try: app.unregisterCustomEvent(CUSTOM_EVENT_ID)
        except: pass
        
        custom_event = app.registerCustomEvent(CUSTOM_EVENT_ID)
        on_custom_event = SuccessTriggerHandler()
        custom_event.add(on_custom_event)
        handlers.append(on_custom_event)

        cmdDef = ui.commandDefinitions.itemById('ModingeFusion360_CSV_Manager')
        if cmdDef: cmdDef.deleteMe()
        cmdDef = ui.commandDefinitions.addButtonDefinition('ModingeFusion360_CSV_Manager', 'ModingeFusion360-ParameterManagerCSVExporter', '')
        
        on_created = MainCommandCreatedEventHandler()
        cmdDef.commandCreated.add(on_created)
        handlers.append(on_created)
        
        cmdDef.execute()
    except: 
        ui.messageBox(traceback.format_exc())

def stop(context):
    try:
        app.unregisterCustomEvent(CUSTOM_EVENT_ID)
        cmdDef1 = ui.commandDefinitions.itemById('ModingeFusion360_CSV_Manager')
        if cmdDef1: cmdDef1.deleteMe()
        cmdDef2 = ui.commandDefinitions.itemById('ModingeFusion360_Success')
        if cmdDef2: cmdDef2.deleteMe()
    except: pass