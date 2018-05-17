#Author-Wayne Brill, Luke Dempsey
#Description-Allows you to select a CSV (comma seperated values) file and then edits existing Attributes. Also allows you to write parameters to a file

import adsk.core, adsk.fusion, traceback
from . import utils


# global set of event handlers to keep them referenced for the duration of the command
handlers = []

class ParameterIOCommand(object):
    def __init__(self):
        self.exportOnlyUserParams = False
        self.importParams = True        #otherwise export
        self.commandId = 'ParamsFromCSV'
        self.workspaceToUse = 'FusionSolidEnvironment'
        self.panelToUse = 'SolidModifyPanel'
        self.app = adsk.core.Application.get()
        self.ui = self.app.userInterface
        self.design = self.app.activeProduct
        self.handlers = utils.HandlerHelper()

    def resetToDefault(self):
        self.exportOnlyUserParams = True
        self.importParams = True

    def commandDefinitionById(self, id):
        if not id:
            self.ui.messageBox('commandDefinition id is not specified')
            return None
        commandDefinitions_ = self.ui.commandDefinitions
        commandDefinition_ = commandDefinitions_.itemById(id)
        return commandDefinition_

    def commandControlByIdForQAT(self, id):
        app = adsk.core.Application.get()
        ui = app.userInterface
        if not id:
            ui.messageBox('commandControl id is not specified')
            return None
        toolbars_ = ui.toolbars
        toolbarQAT_ = toolbars_.itemById('QAT')
        toolbarControls_ = toolbarQAT_.controls
        toolbarControl_ = toolbarControls_.itemById(id)
        return toolbarControl_

    def commandControlByIdForPanel(self, id):
        app = adsk.core.Application.get()
        ui = app.userInterface
        if not id:
            ui.messageBox('commandControl id is not specified')
            return None
        workspaces_ = ui.workspaces
        modelingWorkspace_ = workspaces_.itemById(self.workspaceToUse)
        toolbarPanels_ = modelingWorkspace_.toolbarPanels
        toolbarPanel_ = toolbarPanels_.itemById(self.panelToUse)
        toolbarControls_ = toolbarPanel_.controls
        toolbarControl_ = toolbarControls_.itemById(id)
        return toolbarControl_

    def addButton(self):
        # clean up any crashed instances of the button if existing
        try:
            self.removeButton()
            self.resetToDefault()
        except:
            pass


        # add add-in to UI
        buttonParameterIO = self.ui.commandDefinitions.addButtonDefinition(
            self.commandId, 'ParameterIO (Beta)', 'Import/Exports Parameters', 'resources/command')

        buttonParameterIO.commandCreated.add(self.handlers.make_handler(adsk.core.CommandCreatedEventHandler,
                                                                    self.onCreate))

        createPanel = self.ui.allToolbarPanels.itemById('SolidModifyPanel')
        buttonControl = createPanel.controls.addCommand(buttonParameterIO, 'parameterIOBtn')

        # Make the button available in the panel.
        buttonControl.isPromotedByDefault = True
        buttonControl.isPromoted = True

    def removeButton(self):
        cmdDef = self.ui.commandDefinitions.itemById(self.commandId)
        if cmdDef:
            cmdDef.deleteMe()
        createPanel = self.ui.allToolbarPanels.itemById('SolidModifyPanel')
        cntrl = createPanel.controls.itemById(self.commandId)
        if cntrl:
            cntrl.deleteMe()

    def onCreate(self, args):
        inputs = args.command.commandInputs
        self.resetToDefault()
        args.command.setDialogInitialSize(425, 475)
        args.command.setDialogMinimumSize(425, 475)

        inputs.addBoolValueInput("importParams",
                                       "Import Parameters (untick to export)",
                                       True, "", self.importParams)

        userparams = inputs.addBoolValueInput("exportOnlyUserParams",
                                       "Export Only User Parameters",
                                       True, "", self.exportOnlyUserParams)
        userparams.isEnabled = not(self.importParams)
        
        inputChanged = self.inputChangedEvent()
        args.command.inputChanged.add(inputChanged)
        handlers.append(inputChanged)

        # Add handlers to this command.
        args.command.execute.add(self.handlers.make_handler(adsk.core.CommandEventHandler, self.onExecute))
        #args.command.validateInputs.add(
        #    self.handlers.make_handler(adsk.core.ValidateInputsEventHandler, self.onValidate))
        #args.command.inputChanged.add(
        #    self.handlers.make_handler(adsk.core.InputChangedEventHandler, self.onInputChanged))

    class inputChangedEvent(adsk.core.InputChangedEventHandler):
        def __init__(self):
            super().__init__()
        def notify(self, args):
            try:
                eventArgs = adsk.core.InputChangedEventArgs.cast(args)
                
                changedInput = eventArgs.input
                
                userparam = None
                if changedInput.id == 'importParams':
                    inputs = eventArgs.firingEvent.sender.commandInputs
                    userparam = inputs.itemById('exportOnlyUserParams')
                
                # Change the visibility of the scale value input.
                if userparam:
                    if changedInput.value == True:
                        userparam.isEnabled = False
                    else:
                        userparam.isEnabled = True
    
            except:
                if self.ui:
                    self.ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

    def onExecute(self, args):
        self.parseInputs(args.firingEvent.sender.commandInputs)
        self.updateParamsFromCSV()
        
    def parseInputs(self, inputs):
        inputs = {inp.id: inp for inp in inputs}
        self.exportOnlyUserParams = inputs['exportOnlyUserParams'].value
        self.importParams = inputs['importParams'].value

    class CommandCreatedEventHandlerQAT(adsk.core.CommandCreatedEventHandler):
        def __init__(self):
            super().__init__()
        def notify(self, args):
            try:
                command = args.command
                onExecute = self.CommandExecuteHandler()
                command.execute.add(onExecute)
                # keep the handler referenced beyond this function
                handlers.append(onExecute)

            except:
                self.ui.messageBox('QAT command created failed:\n{}'.format(traceback.format_exc()))


    def destroyObject(self, uiObj, tobeDeleteObj):
        if uiObj and tobeDeleteObj:
            if tobeDeleteObj.isValid:
                tobeDeleteObj.deleteMe()
            else:
                uiObj.messageBox('tobeDeleteObj is not a valid object')

    def run(self):
        try:
            commandName = 'Import/Export Parameters (CSV)'
            commandDescription = 'Import parameters from or export them to a CSV (Comma Separated Values) file'
            commandResources = './resources/command'

            commandDefinitions_ = self.ui.commandDefinitions
            
            #buttonDogbone = self.ui.commandDefinitions.addButtonDefinition(
            #self.COMMAND_ID, 'Dogbone', 'Creates a dogbone at the corner of two lines/edges', 'Resources')

            # check if we have the command definition
            commandDefinition_ = commandDefinitions_.itemById(self.commandId)
            if not commandDefinition_:
                commandDefinition_ = commandDefinitions_.addButtonDefinition(self.commandId, commandName, commandName, commandResources)
                commandDefinition_.tooltipDescription = commandDescription       

            onCommandCreated = self.CommandCreatedEventHandlerPanel()
            commandDefinition_.commandCreated.add(onCommandCreated)
            # keep the handler referenced beyond this function
            handlers.append(onCommandCreated)

            # add a command on create panel in modeling workspace
            workspaces_ = self.ui.workspaces
            modelingWorkspace_ = workspaces_.itemById(self.workspaceToUse)
            toolbarPanels_ = modelingWorkspace_.toolbarPanels
            toolbarPanel_ = toolbarPanels_.itemById(self.panelToUse) 
            toolbarControlsPanel_ = toolbarPanel_.controls
            toolbarControlPanel_ = toolbarControlsPanel_.itemById(self.commandId)
            if not toolbarControlPanel_:
                toolbarControlPanel_ = toolbarControlsPanel_.addCommand(commandDefinition_, self.commandId)
                toolbarControlPanel_.isVisible = True
                #ui.messageBox('A CSV command is successfully added to the create panel in modeling workspace')

        except:
            if self.ui:
               self.ui.messageBox('AddIn Start Failed:\n{}'.format(traceback.format_exc()))

    def stop(self):
        try:
            objArray = []

            commandControlQAT_ = self.commandControlByIdForQAT(self.commandId)
            if commandControlQAT_:
                objArray.append(commandControlQAT_)

            commandControlPanel_ = self.commandControlByIdForPanel(self.commandId)
            if commandControlPanel_:
                objArray.append(commandControlPanel_)
                
            commandDefinition_ = self.commandDefinitionById(self.commandId)
            if commandDefinition_:
                objArray.append(commandDefinition_)

            for obj in objArray:
                self.destroyObject(self.ui, obj)

        except:
            if self.ui:
                self.ui.messageBox('AddIn Stop Failed:\n{}'.format(traceback.format_exc()))

    def updateParamsFromCSV(self):
         
        try:
             #Ask if reading or writing parameters
            
            fileDialog = self.ui.createFileDialog()
            fileDialog.isMultiSelectEnabled = False
            fileDialog.title = "Get the file to read from or the file to save the parameters to"
            fileDialog.filter = 'Text files (*.csv)'
            fileDialog.filterIndex = 0
            if self.importParams:
                dialogResult = fileDialog.showOpen()
            else:
                dialogResult = fileDialog.showSave()
                 
            if dialogResult == adsk.core.DialogResults.DialogOK:
                filename = fileDialog.filename
            else:
                return

            #if readParameters is true read the parameters from a file
            if self.importParams:
                self.readTheParameters(filename)
            else:
                self.writeTheParameters(filename)

        except:
            if self.ui:
                self.ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

                
    def writeTheParameters(self, theFileName):
        app = adsk.core.Application.get()
        design = app.activeProduct
        result = ""

        if self.exportOnlyUserParams:
            paramsToExport = design.userParameters
        else:
            paramsToExport = design.allParameters  

        for _param in paramsToExport:
            result = result + _param.name +  "," + _param.unit +  "," + _param.expression + "," + _param.comment + "\n"

        outputFile = open(theFileName, 'w')
        outputFile.writelines(result)
        outputFile.close()
        
        #get the name of the file without the path    
        pathsInTheFileName = theFileName.split("/")
        self.ui.messageBox('Parameters written to ' + pathsInTheFileName[-1])   
       
    def readTheParameters(self, theFileName):
        
        try:
            paramsList = []
            for oParam in self.design.allParameters:
                paramsList.append(oParam.name)           
            
            # Read the csv file.
            csvFile = open(theFileName)
            for line in csvFile:
                # Get the values from the csv file.
                # remove end line characters
                line = line.rstrip('\n\r')
                valsInTheLine = line.split(',')
                nameOfParam = valsInTheLine[0]
                unitOfParam = valsInTheLine[1]
                expressionOfParam = valsInTheLine[2]
                # userParameters.add does not like empty string as comment
                # so we make it a space
                commentOfParam = ' '
                # comment might be missing
                if len(valsInTheLine) > 3:
                    # if it's not an empty string    
                    if valsInTheLine[3] != '':
                        commentOfParam = valsInTheLine[3] 
                    
                # if the name of the paremeter is not an existing parameter add it
                if nameOfParam not in paramsList:
                    valInput_Param = adsk.core.ValueInput.createByString(expressionOfParam) 
                    self.design.userParameters.add(nameOfParam, valInput_Param, unitOfParam, commentOfParam)
                # update the values of existing parameters            
                else:
                    paramInModel = self.design.allParameters.itemByName(nameOfParam)
                    paramInModel.unit = unitOfParam
                    paramInModel.expression = expressionOfParam
                    paramInModel.comment = commentOfParam
            self.ui.messageBox('Finished reading and updating parameters')
        except:
            if self.ui:
                self.ui.messageBox('AddIn Stop Failed:\n{}'.format(traceback.format_exc()))

paramIO = ParameterIOCommand()

def run(context):
    try:
        paramIO.addButton()
    except:
        utils.messageBox(traceback.format_exc())


def stop(context):
    try:
        paramIO.removeButton()
    except:
        utils.messageBox(traceback.format_exc())
