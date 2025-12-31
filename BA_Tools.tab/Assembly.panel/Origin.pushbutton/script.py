"""
Revit API Script: Rotate Assembly Instances by Specified Axis and Angle
Inputs:
- Assembly instances
- Rotation axis (X, Y, or Z)
- Rotation angle in degrees
"""

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
import math

# Get the current Revit application and document
__revit__ = __revit__
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document

# ============================================
# INPUTS - Modify these as needed
# ============================================

# INPUT 1: Assembly instances (list of AssemblyInstance objects)
# You can get these by selection or filtering
assembly_instances = []  # Replace with your assembly instances

# INPUT 2: Rotation axis - "X", "Y", or "Z"
rotation_axis = "Z"  # Options: "X", "Y", "Z"

# INPUT 3: Rotation angle in degrees
rotation_angle_degrees = 90.0  # Angle in degrees

# ============================================
# FUNCTIONS
# ============================================

def get_rotation_axis_vector(assembly, axis_name):
    """
    Get the rotation axis line based on assembly origin and specified axis
    
    Args:
        assembly: AssemblyInstance object
        axis_name: String - "X", "Y", or "Z"
    
    Returns:
        Line object representing the rotation axis
    """
    # Get assembly origin point
    assembly_origin = assembly.GetTransform().Origin
    
    # Define axis direction vector
    if axis_name.upper() == "X":
        axis_vector = XYZ.BasisX
    elif axis_name.upper() == "Y":
        axis_vector = XYZ.BasisY
    elif axis_name.upper() == "Z":
        axis_vector = XYZ.BasisZ
    else:
        raise ValueError("Axis must be 'X', 'Y', or 'Z'")
    
    # Create axis line through assembly origin
    point1 = assembly_origin
    point2 = assembly_origin + axis_vector
    axis_line = Line.CreateBound(point1, point2)
    
    return axis_line

def rotate_assembly(doc, assembly, axis_name, angle_degrees):
    """
    Rotate an assembly instance around specified axis by given angle
    
    Args:
        doc: Revit document
        assembly: AssemblyInstance object
        axis_name: String - "X", "Y", or "Z"
        angle_degrees: Float - rotation angle in degrees
    
    Returns:
        Boolean - True if successful, False otherwise
    """
    try:
        # Convert angle from degrees to radians
        angle_radians = math.radians(angle_degrees)
        
        # Get rotation axis line
        axis_line = get_rotation_axis_vector(assembly, axis_name)
        
        # Perform rotation
        ElementTransformUtils.RotateElement(doc, assembly.Id, axis_line, angle_radians)
        
        return True
        
    except Exception as e:
        print("Error rotating assembly {0}: {1}".format(assembly.Id, str(e)))
        return False

def select_assemblies(uidoc):
    """
    Prompt user to select assembly instances
    
    Returns:
        List of AssemblyInstance objects
    """
    assemblies = []
    try:
        selection = uidoc.Selection
        references = selection.PickObjects(ObjectType.Element, "Select assembly instances (ESC when done)")
        
        for ref in references:
            element = doc.GetElement(ref.ElementId)
            if isinstance(element, AssemblyInstance):
                assemblies.append(element)
            else:
                print("Element {0} is not an assembly instance".format(element.Id))
        
        return assemblies
        
    except:
        return assemblies

# ============================================
# MAIN EXECUTION
# ============================================

try:
    # If assembly_instances list is empty, prompt user to select
    if len(assembly_instances) == 0:
        TaskDialog.Show("Select Assemblies", 
                       "Please select assembly instances to rotate\n" +
                       "Press ESC when done selecting")
        assembly_instances = select_assemblies(uidoc)
    
    if len(assembly_instances) == 0:
        TaskDialog.Show("No Selection", "No assembly instances selected")
    else:
        # Start transaction
        transaction = Transaction(doc, "Rotate Assembly Instances")
        transaction.Start()
        
        success_count = 0
        failed_count = 0
        
        # Rotate each assembly
        for assembly in assembly_instances:
            result = rotate_assembly(doc, assembly, rotation_axis, rotation_angle_degrees)
            
            if result:
                success_count += 1
            else:
                failed_count += 1
        
        # Commit transaction
        transaction.Commit()
        
        # Show results
        result_message = "Rotation Complete!\n\n"
        result_message += "Axis: {0}\n".format(rotation_axis)
        result_message += "Angle: {0}°\n\n".format(rotation_angle_degrees)
        result_message += "Successfully rotated: {0} assemblies\n".format(success_count)
        
        if failed_count > 0:
            result_message += "Failed: {0} assemblies".format(failed_count)
        
        TaskDialog.Show("Results", result_message)

except Exception as ex:
    TaskDialog.Show("Error", "An error occurred:\n" + str(ex))

