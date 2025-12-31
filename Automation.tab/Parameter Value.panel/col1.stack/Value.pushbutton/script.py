
"""
Set Mark Values for Elements
Select category, view, filter by type, and assign mark values
"""
__title__ = 'Set Mark\nValues'
__doc__ = 'Assign mark values to filtered elements by type in a selected view'
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script
doc = revit.doc
uidoc = revit.uidoc
# ==============================================================================
# STEP 1: SELECT CATEGORY
# ==============================================================================
category_options = [
    'Walls',
    'Doors',
    'Windows',
    'Structural Framing',
    'Structural Columns',
    'Floors',
    'Furniture',
    'Generic Models'
]
category_map = {
    'Walls': BuiltInCategory.OST_Walls,
    'Doors': BuiltInCategory.OST_Doors,
    'Windows': BuiltInCategory.OST_Windows,
    'Structural Framing': BuiltInCategory.OST_StructuralFraming,
    'Structural Columns': BuiltInCategory.OST_StructuralColumns,
    'Floors': BuiltInCategory.OST_Floors,
    'Furniture': BuiltInCategory.OST_Furniture,
    'Generic Models': BuiltInCategory.OST_GenericModel
}
selected_category_name = forms.SelectFromList.show(
    category_options,
    title='Step 1: Select Category',
    button_name='Select'
)
if not selected_category_name:
    script.exit()
selected_category = category_map[selected_category_name]
# ==============================================================================
# STEP 2: SELECT VIEW
# ==============================================================================
all_views = FilteredElementCollector(doc)\
    .OfClass(View)\
    .WhereElementIsNotElementType()\
    .ToElements()
valid_views = [v for v in all_views if not v.IsTemplate and v.CanBePrinted]
if not valid_views:
    forms.alert('No valid views found', exitscript=True)
view_dict = {}
for view in valid_views:
    view_dict[view.Name] = view
selected_view_name = forms.SelectFromList.show(
    sorted(view_dict.keys()),
    title='Step 2: Select View',
    button_name='Select'
)
if not selected_view_name:
    script.exit()
selected_view = view_dict[selected_view_name]
# ==============================================================================
# STEP 3: COLLECT ALL ELEMENTS IN VIEW
# ==============================================================================
all_elements = FilteredElementCollector(doc, selected_view.Id)\
    .OfCategory(selected_category)\
    .WhereElementIsNotElementType()\
    .ToElements()
if not all_elements:
    forms.alert(
        'No {} found in view "{}"'.format(selected_category_name, selected_view.Name),
        exitscript=True
    )
print('Found {} {} in view "{}"'.format(
    len(all_elements), 
    selected_category_name, 
    selected_view.Name
))
# ==============================================================================
# STEP 4: GROUP ELEMENTS BY TYPE
# ==============================================================================
type_dict = {}
for elem in all_elements:
    type_id = elem.GetTypeId()
    elem_type = doc.GetElement(type_id)
    
    if elem_type:
        type_name = elem_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        
        if type_name not in type_dict:
            type_dict[type_name] = []
        
        type_dict[type_name].append(elem)
# ==============================================================================
# STEP 5: SELECT TYPE TO FILTER
# ==============================================================================
type_options = []
for type_name, elems in sorted(type_dict.items()):
    type_options.append('{} ({} elements)'.format(type_name, len(elems)))
selected_type_option = forms.SelectFromList.show(
    type_options,
    title='Step 3: Select Element Type',
    button_name='Select',
    multiselect=False
)
if not selected_type_option:
    script.exit()
# Extract type name from selection (remove count)
selected_type_name = selected_type_option.split(' (')[0]
# Get filtered elements
filtered_elements = type_dict[selected_type_name]
print('\nFiltered to {} elements of type "{}"'.format(
    len(filtered_elements), 
    selected_type_name
))
# ==============================================================================
# STEP 6: SHOW CURRENT MARK VALUES
# ==============================================================================
print('\nCurrent Mark Values:')
print('-' * 50)
for i, elem in enumerate(filtered_elements, 1):
    mark_param = elem.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
    current_mark = ''
    
    if mark_param and mark_param.HasValue:
        current_mark = mark_param.AsString() or ''
    
    print('{}. Element ID: {} | Current Mark: "{}"'.format(
        i, 
        elem.Id, 
        current_mark if current_mark else '(Empty)'
    ))
print('-' * 50)
# ==============================================================================
# STEP 7: ASK FOR MARK VALUE INPUT METHOD
# ==============================================================================
input_method = forms.CommandSwitchWindow.show(
    [
        'Single Value for All',
        'Sequential Numbers',
        'Sequential with Prefix',
        'Comma-Separated Values'
    ],
    message='Step 4: Choose how to assign mark values:'
)
if not input_method:
    script.exit()
# ==============================================================================
# STEP 8: GET MARK VALUES FROM USER
# ==============================================================================
new_marks = []
if input_method == 'Single Value for All':
    # Ask for one value to apply to all elements
    value = forms.ask_for_string(
        prompt='Enter mark value for all {} elements:'.format(len(filtered_elements)),
        title='Enter Mark Value',
        default=''
    )
    
    if value is None:
        script.exit()
    
    new_marks = [value] * len(filtered_elements)
elif input_method == 'Sequential Numbers':
    # Ask for starting number
    start_value = forms.ask_for_string(
        prompt='Enter starting number:',
        title='Starting Number',
        default='1'
    )
    
    if not start_value:
        script.exit()
    
    try:
        start_num = int(start_value)
        new_marks = [str(start_num + i) for i in range(len(filtered_elements))]
    except:
        forms.alert('Invalid number entered', exitscript=True)
elif input_method == 'Sequential with Prefix':
    # Ask for prefix
    prefix = forms.ask_for_string(
        prompt='Enter prefix (e.g., W, WALL, A):',
        title='Enter Prefix',
        default=''
    )
    
    if prefix is None:
        script.exit()
    
    # Ask for starting number
    start_value = forms.ask_for_string(
        prompt='Enter starting number:',
        title='Starting Number',
        default='1'
    )
    
    if not start_value:
        script.exit()
    
    try:
        start_num = int(start_value)
        new_marks = ['{}-{}'.format(prefix, start_num + i) for i in range(len(filtered_elements))]
    except:
        forms.alert('Invalid number entered', exitscript=True)
elif input_method == 'Comma-Separated Values':
    # Show info and ask for comma-separated values
    prompt_text = 'Enter {} mark values separated by commas.\n\n'.format(len(filtered_elements))
    prompt_text += 'Example: W1, W2, W3, W4\n\n'
    prompt_text += 'Elements to be marked:\n'
    
    for i in range(min(5, len(filtered_elements))):
        mark_param = filtered_elements[i].get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        current = ''
        if mark_param and mark_param.HasValue:
            current = mark_param.AsString() or ''
        
        prompt_text += '{}. ID: {} (Current: {})\n'.format(
            i + 1,
            filtered_elements[i].Id,
            current if current else 'Empty'
        )
    
    if len(filtered_elements) > 5:
        prompt_text += '... and {} more'.format(len(filtered_elements) - 5)
    
    values_input = forms.ask_for_string(
        prompt=prompt_text,
        title='Enter Mark Values'
    )
    
    if not values_input:
        script.exit()
    
    # Parse input
    entered_values = [v.strip() for v in values_input.split(',')]
    
    if len(entered_values) != len(filtered_elements):
        forms.alert(
            'Mismatch!\n\nYou entered {} values but there are {} elements.\n\nPlease enter exactly {} comma-separated values.'.format(
                len(entered_values),
                len(filtered_elements),
                len(filtered_elements)
            ),
            exitscript=True
        )
    
    new_marks = entered_values
# ==============================================================================
# STEP 9: APPLY MARK VALUES TO ELEMENTS
# ==============================================================================
success_count = 0
error_count = 0
results = []
t = Transaction(doc, 'Set Mark Values')
t.Start()
try:
    for elem, new_mark in zip(filtered_elements, new_marks):
        try:
            mark_param = elem.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
            
            if mark_param and not mark_param.IsReadOnly:
                # Get old value for reporting
                old_mark = ''
                if mark_param.HasValue:
                    old_mark = mark_param.AsString() or ''
                
                # Set new value
                mark_param.Set(new_mark)
                
                results.append({
                    'id': elem.Id,
                    'old': old_mark,
                    'new': new_mark,
                    'success': True
                })
                
                success_count += 1
            else:
                results.append({
                    'id': elem.Id,
                    'old': '',
                    'new': new_mark,
                    'success': False,
                    'error': 'Read-only parameter'
                })
                error_count += 1
                
        except Exception as e:
            results.append({
                'id': elem.Id,
                'old': '',
                'new': new_mark,
                'success': False,
                'error': str(e)
            })
            error_count += 1
    
    t.Commit()
    
except Exception as e:
    t.RollBack()
    forms.alert('Transaction failed: {}'.format(str(e)), exitscript=True)
# ==============================================================================
# STEP 10: DISPLAY RESULTS
# ==============================================================================
print('\n' + '=' * 70)
print('MARK VALUES UPDATED')
print('=' * 70)
print('Category: {}'.format(selected_category_name))
print('View: {}'.format(selected_view.Name))
print('Type: {}'.format(selected_type_name))
print('-' * 70)
print('Success: {} elements'.format(success_count))
print('Errors: {} elements'.format(error_count))
print('=' * 70)
print('\nDetailed Results:')
for i, result in enumerate(results, 1):
    if result['success']:
        print('{}. ID: {} | "{}" -> "{}"'.format(
            i,
            result['id'],
            result['old'] if result['old'] else '(Empty)',
            result['new']
        ))
    else:
        print('{}. ID: {} | ERROR: {}'.format(
            i,
            result['id'],
            result.get('error', 'Unknown error')
        ))
print('=' * 70)
# Show completion dialog
forms.alert(
    'Mark values updated!\n\n'
    'Success: {} elements\n'
    'Errors: {} elements\n\n'
    'Check output window for details.'.format(success_count, error_count),
    title='Completed'
)

