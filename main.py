import os
import subprocess

def create_shortcut(name, target_path, icon_path=None):
    # Create a temporary VBS script file
    vbs_script = f"""
    Set oWS = WScript.CreateObject("WScript.Shell")
    sLinkFile = "{name}.lnk"
    sTargetPath = "{target_path}"
    sIconPath = "{icon_path}" if "{icon_path}" else sTargetPath
    Set oLink = oWS.CreateShortcut(sLinkFile)
    oLink.TargetPath = sTargetPath
    oLink.IconLocation = sIconPath
    oLink.Save
    """

    script_file = os.path.join(os.getcwd(), 'temp.vbs')

    with open(script_file, 'w') as file:
        file.write(vbs_script)

    # Execute the VBS script to create the shortcut
    subprocess.call(['cscript', script_file], shell=True)
    
    # Remove the temporary script file
    os.remove(script_file)

# Provide the name, target path, and optional icon path for the shortcut
shortcut_name = "My Shortcut"
target_path = "C:\\path\\to\\target\\file.exe"
icon_path = "C:\\path\\to\\icon.ico"  # Optional: Comment out or set to None if no icon needed

# Create the shortcut
create_shortcut(shortcut_name, target_path, icon_path)
