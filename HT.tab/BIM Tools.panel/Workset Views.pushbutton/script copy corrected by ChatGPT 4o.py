import clr
import System
from System.Collections.Generic import List
from pyrevit import revit, DB, script, forms, framework

# Get the current document
doc = revit.doc

class AllViewTemplates(forms.TemplateListItem):
    @property
    def name(self):
        return self.item.Name

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
        visible_buttons = [b for b in self.button_list.Children if b.Visibility == framework.Windows.Visibility.Visible]
        return visible_buttons[0] if len(visible_buttons) == 1 else next((b for b in visible_buttons if b.IsFocused), None)

    def handle_click(self, sender, args):
        self.Close()

    def handle_input_key(self, sender, args):
        if args.Key == framework.Windows.Input.Key.Escape:
            self.Close()
        elif args.Key == framework.Windows.Input.Key.Enter:
            self.process_option(self._get_active_button(), None)

    def process_option(self, sender, args):
        self.Close()
        if sender:
            self._setup_response(response=sender.Content)


# Create a FilteredWorksetCollector to get all Worksets in the document
worksets_collection = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset)

# Check if there are any Worksets
if len(worksets_collection) == 1 and worksets_collection[0].Name == "Workset1":
    forms.alert("No Worksets found in the project.", title="Workset Info")
else:
    worksetsDict = {workset.Name: workset for workset in worksets_collection}
    collector = DB.FilteredElementCollector(doc).OfClass(DB.View3D)
    
    viewTemplates3D = [v for v in collector if v.IsTemplate]
    existingWorkset3DViews = [v for v in collector if not v.IsTemplate and v.Name in worksetsDict]
    
    for v in existingWorkset3DViews:
        worksetsDict.pop(v.Name, None)

    viewTemplates3D.sort(key=lambda obj: obj.Name)

    class NoneOption:
        def __init__(self, name):
            self.Name = name

    viewTemplates3D.append(NoneOption('<None>'))

    viewType = next(x for x in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType) if x.ViewFamily == DB.ViewFamily.ThreeDimensional)

    new3DViews = []
    new3DViewsNames = []
    allViews = []

    def create3DViewsForEachWorkset():
        new3DViews.extend(
            DB.View3D.CreateIsometric(doc, viewType.Id)
            for worksetName in worksetsDict
        )
        for view in new3DViews:
            try:
                view.Name = worksetsDict[view.Name].Name
                new3DViewsNames.append(view.Name)
            except:
                pass
        allViews.extend(new3DViews + existingWorkset3DViews)
        return allViews

    def applyViewTemplate(viewTemplate):
        view_template_not_controlled_settings = viewTemplate.GetNonControlledTemplateParameterIds()
        worksetsNonControlled = '-1006968' in map(str, view_template_not_controlled_settings)

        if worksetsNonControlled:
            for v in create3DViewsForEachWorkset():
                v.ViewTemplateId = viewTemplate.Id
            return True
        else:
            selected_option = SelectOverrideOpt.show(
                ['Yes', 'Cancel & Exit'],
                response='Yes',
                message='Selected View Template setting for Workset V/G Overrides\nmust be unchecked to create Workset Views.\n\nWould you like to uncheck it?\n',
                title='Workset V/G Overrides setting',
            )
            if selected_option == 'Yes':
                view_template_not_controlled_settings.Add(DB.ElementId(-1006968))
                viewTemplate.SetNonControlledTemplateParameterIds(view_template_not_controlled_settings)
                for v in create3DViewsForEachWorkset():
                    v.ViewTemplateId = viewTemplate.Id
                return True
            return False

    return_viewTemplate = forms.SelectFromList.show(
        [AllViewTemplates(x) for x in viewTemplates3D],
        title='Select a View Template to use for Workset 3D views',
        width=470,
        button_name='Create views',
        multiselect=False
    )

    visibilities = list(System.Enum.GetValues(DB.WorksetVisibility))
    visible = visibilities[0]
    hidden = visibilities[1]

    alertTitle = 'Explanation'
    alertMessage = 'No Workset Views were created. View Template setting for Workset V/G Overrides must be unchecked to create Workset Views.'

    if return_viewTemplate:
        with revit.Transaction('Create Views for each Workset'):
            if return_viewTemplate.Name != '<None>':
                if not applyViewTemplate(return_viewTemplate):
                    forms.alert(alertMessage, alertTitle)
            else:
                default3DViewTemplateId = viewType.DefaultTemplateId
                if default3DViewTemplateId:
                    if viewType.get_Parameter(DB.BuiltInParameter.ASSIGN_TEMPLATE_ON_VIEW_CREATION).AsInteger() == 1:
                        default3DViewTemplate = doc.GetElement(default3DViewTemplateId)
                        if not applyViewTemplate(default3DViewTemplate):
                            forms.alert(alertMessage, alertTitle)
                    else:
                        create3DViewsForEachWorkset()
                else:
                    create3DViewsForEachWorkset()

            for workset in worksets_collection:
                for v in allViews:
                    try:
                        visibility = visible if workset.Name == v.Name else hidden
                        v.SetWorksetVisibility(workset.Id, visibility)
                    except Exception as del_err:
                        logger = script.get_logger()
                        logger.error(f'Error applying workset visibility: {workset.Name} | {del_err}')
                        forms.alert(f'Error applying workset visibility: {workset.Name} | {del_err}')

        final_message = ''
        if new3DViewsNames:
            final_message = 'New 3D Views created:\n- ' + "\n- ".join(new3DViewsNames)
        if existingWorkset3DViews:
            num_updated = len(existingWorkset3DViews)
            update_message = f'\n\n{num_updated} existing Workset 3D View{"s were" if num_updated > 1 else " was"} updated.'
            final_message += update_message

        if final_message:
            forms.alert(final_message, title="New 3D Workset Views")
