# Author: Pawel Block
# Company: Haworth Tompkins Ltd
# Date: 2024-05-14
# Version: 1.0.3
# Description: This script will create a 3D view for each workset and will update existing Workset views. It allows to specify a View Template which can direct new views to specific Project Browser folders.
# Tested with: Revit +2022
# Requirements: pyRevit add-in

import System
# Import pyRevit modules
from pyrevit import revit, DB, script, forms, framework

# Get the current document
doc = revit.doc

class AllViewTemplates(forms.TemplateListItem):
    @property
    def name(self):
        return self.item.Name

# Below is a simplified code from pyRevit class pyrevit.forms.CommandSwitchWindow
class SelectOverrideOpt(forms.TemplateUserInputWindow):
    xaml_source = script.get_bundle_file('options.xaml')

    def _setup(self, **kwargs):
        message = kwargs.get('message', 'Pick a command option:')
        self.message_label.Content = message

        for option in self._context:
            my_button = framework.Controls.Button()
            my_button.Content = option
            my_button.Click += self.process_option
            self.button_list.Children.Add(my_button)
        self._setup_response()

    def _setup_response(self, response=None):
        self.response = response

    def _get_active_button(self):
        buttons = []
        for button in self.button_list.Children:
            if button.Visibility == framework.Windows.Visibility.Visible:
                buttons.append(button)
        if len(buttons) == 1:
            return buttons[0]
        else:
            for x in buttons:
                if x.IsFocused:
                    return x

    def handle_click(self, sender, args):
        """Handle mouse click."""
        self.Close()

    def handle_input_key(self, sender, args):
        """Handle keyboard inputs."""
        if args.Key == framework.Windows.Input.Key.Escape:
            self.Close()
        elif args.Key == framework.Windows.Input.Key.Enter:
            self.process_option(self._get_active_button(), None)

    def process_option(self, sender, args):
        """Handle click on command option button."""
        self.Close()
        if sender:
            self._setup_response(response=sender.Content)

# Create a FilteredWorksetCollector to get all Worksets in the document
worksets_collection = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset)
worksets = list(worksets_collection)

# Checks if there are any Worksets
if len(worksets) == 1 and worksets[0].Name == "Workset1":
    forms.alert("No Worksets found in the project.", title="Workset Info")
else:
    # Creates a list to store Workset names
    worksetsDict = {}

    # Iterates over each Workset and get its name
    for workset in worksets:
        # This will sore worksets with no corresponding view
        worksetsDict[workset.Name] = workset

    collector = DB.FilteredElementCollector(doc).OfClass(DB.View3D)
    viewTemplates3D = []
    existingWorkset3DViews = []

    for v in collector:
        if v.IsTemplate == True:
            viewTemplates3D.append(v)
        else:
            if v.Name in worksetsDict:
                worksetsDict.pop(v.Name)
                existingWorkset3DViews.append(v)

    # Sorts a list of list elements alphabetically by name.
    viewTemplates3D.sort(key=lambda obj: obj.Name)

    # Creates an object to add to the list of View Templates representing the None
    class NoneOption():
            def __init__(self, name):
                self.Name = name
    # Adds a None option to the list of 3D View Templates
    viewTemplates3D.append(NoneOption('<None>'))

    #Gets 3D View ViewFamilyType
    viewType = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType).ToElements().Find(lambda x : x.ViewFamily == DB.ViewFamily.ThreeDimensional)

    # Creates new 3D views for each workset missing a view
    def create3DViewsForEachWorkset():
        for worksetName in worksetsDict:
            view = DB.View3D.CreateIsometric(doc, viewType.Id)
            try:view.Name = worksetName
            except:pass
            new3DViews.append(view)
            new3DViewsNames.append(view.Name)

        # Joins new 3D views to the existing ones
        allViews.extend(new3DViews + existingWorkset3DViews)
        return allViews

    def applyViewTemplate(viewTemplate):
        # Gets the non-controlled settings from the View Template
        view_template_not_controlled_settings = viewTemplate.GetNonControlledTemplateParameterIds()
        worksetsNonControlled = False

        # Iterate over each setting ID to find the V/G Overrides Workset setting
        # V/G Overrides Worksets name and ID is fixed: VIS_GRAPHICS_WORKSETS	-1006968
        for settingId in view_template_not_controlled_settings:
            # Check if the setting or V/G Overrides Worksets is not set
            if settingId.ToString() == '-1006968':
                worksetsNonControlled = True

        if worksetsNonControlled == True:
            for v in create3DViewsForEachWorkset():
                v.ViewTemplateId = viewTemplate.Id
            return True
        else:
            opts = ['Yes', 'Cancel & Exit']
            selected_option = SelectOverrideOpt.show(
                opts,
                response = 'Yes',
                message='Selected View Template setting for Workset V/G Overrides\nmust be unchecked to create Workset Views.\n\nWould you like to uncheck it?\n',
            )
            if selected_option == 'Yes':
                view_template_not_controlled_settings.Add(DB.ElementId(-1006968))
                # This sets all other parameters not included in the set to be controlled by the View Template
                viewTemplate.SetNonControlledTemplateParameterIds(view_template_not_controlled_settings)
                # Now it is reasonable to apply View Templates to all views
                for v in create3DViewsForEachWorkset():
                    v.ViewTemplateId = viewTemplate.Id
                return True
            return False

    return_viewTemplate = \
        forms.SelectFromList.show(
            [AllViewTemplates(x) for x in viewTemplates3D],
            title='Select a View Template to use for Workset 3D views',
            width=470,
            button_name='Create views',
            multiselect=False
        )

    new3DViews = []
    new3DViewsNames = []
    allViews = []
    # Enumerates the WorksetVisibilities
    visibilities = System.Enum.GetValues(DB.WorksetVisibility)
    visible = visibilities[0] # Visible
    hidden = visibilities[1] # Hidden

    # Common alert message if no Workset Views were created
    alertTitle = 'Explanation'
    alert = 'No Workset Views were created. View Template setting for Workset V/G Overrides must be unchecked to create Workset Views.'

    if return_viewTemplate:
        # Start Transaction
        with revit.Transaction('Create Views for each Workset'):
            # Sets a selected View Template to new and existing Views
            # A selected View Template must allow to to modify workset visibilities
            if return_viewTemplate.Name != '<None>':
                # Sets the workset visibility of the new 3D views
                viewTemplateApplied = applyViewTemplate(return_viewTemplate)
                if not viewTemplateApplied:
                    forms.alert(alert, alertTitle)
            else:
                default3DViewTemplateId = viewType.DefaultTemplateId
                # Checks if default View Template is assigned to new 3D views
                if default3DViewTemplateId:
                    # "New Views are dependent on Template" check if checked (1) or unchecked (0)
                    if viewType.get_Parameter(DB.BuiltInParameter.ASSIGN_TEMPLATE_ON_VIEW_CREATION).AsInteger() == 1:
                        default3DViewTemplate = doc.GetElement(default3DViewTemplateId)
                        viewTemplateApplied = applyViewTemplate(default3DViewTemplate)
                        if not viewTemplateApplied:
                            forms.alert(alert, alertTitle)
                    else:
                        create3DViewsForEachWorkset()
                else:
                    create3DViewsForEachWorkset()

            for workset in worksets_collection:
                try:
                    for v in allViews:
                        if workset.Name == v.Name:
                            v.SetWorksetVisibility( workset.Id, visible )
                        else:
                            v.SetWorksetVisibility( workset.Id, hidden )
                except Exception as del_err:
                    logger = script.get_logger()
                    logger.error('Error applying workset visibility: {} | {}'
                            .format(workset.Name, del_err))
                    forms.alert('Error applying workset visibility: {} | {}'
                        .format(workset.Name, del_err))

        final_message = ''
        if new3DViewsNames:
            # Creates a string of Workset names
            final_message = 'New 3D Views created:\n- '+"\n- ".join(new3DViewsNames)
        if existingWorkset3DViews and new3DViewsNames:
            final_message += '\n\n'
        if existingWorkset3DViews:
            no = str(len(existingWorkset3DViews))
            if no == '1':
                final_message += '1 existing Workset 3D View was updated.'
            else:
                final_message += no+' existing Workset 3D Views were updated.'
        if final_message:
            # Shows the Workset names in a popup message
            forms.alert(final_message, title="New 3D Workset Views")