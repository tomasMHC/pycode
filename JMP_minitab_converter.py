#!/usr/bin/python3
# This does what the JMP File Conversion Assistant from Minitab Support does,
# only better because it also exports Column Properties into metadata files!

import os
import win32com.client
from glob import iglob
from tkinter import filedialog,simpledialog,Tk
from sys import exit

#Folder selection
getroot = Tk()
root_dir = filedialog.askdirectory(initialdir=os.getcwd(),title="Select directory to scan recursively for *.jmp files",master=getroot)
getroot.destroy()
#Ends program if no folder is selected
if root_dir == '':
    print('No folder selected - ending program')
    exit()

files_to_convert = []
print(f"Searching for JMP files in {root_dir}...")
# Use iglob, an iterator, to make sure we can print the filenames as they're found
for fn in iglob("**/*.jmp",root_dir=root_dir,recursive=True):
    print(f"Found file {fn}")
    files_to_convert.append(os.path.join(root_dir,fn))
print("Search complete!")

if not len(files_to_convert):
    print("No files - ending program")
    exit()

# Start up JMP using OLE Automation
print("\nStarting JMP...")
jmp = win32com.client.Dispatch("JMP.Application")
jmp.Visible = True
print(f"JMP running @ {jmp}")

failed_files = []
for fpth in files_to_convert:
    csvtarget = f"{fpth}.csv"
    metatarget = f"{csvtarget}.meta"
    if os.path.isfile(metatarget) and os.path.isfile(csvtarget):
        # Don't redo what's already done
        print(f"{fpth} already converted")
        continue
    print(f"Working on {fpth}...")
    jmpdoc = jmp.OpenDocument(fpth)
    if jmpdoc is None:
        failed_files.append(fpth)
        continue
    if not os.path.isfile(metatarget): # Don't redo what's already done
        jmpdoc.Activate()
        # JSL within Python - "meta" indeed!
        jmp.RunCommand(r"""
        dt = Current Data Table();
        cols = dt << Get Column Names();
        prop_string = ""; /* reinitialize the output string */
        foreach( {col, idx}, cols,
            colname = col << Get Name;
            props = col << Get Column Properties();
            foreach( {prop, propidx}, props,
                head = Head Name(prop);
                if( head == "Set Property",
                    /* 'prop' is a "Set Property" expression, so extract its two arguments and append a row to the output. */
                    prop_string = prop_string || colname || "\!t" || Arg(prop,1) || "\!t" || Char(Arg(prop,2)) || "\!n";
                    , /*else if*/ N Arg(prop) == 1,
                    /* one argument, which is the full formula / property */
                    prop_string = prop_string || colname || "\!t" || head || "\!t" || Char(Arg(prop,1)) || "\!n";
                    , /*else*/
                    prop_string = prop_string || colname || "\!t" || head || "\!t" || Char(prop) || "\!n"; /* just reproduce the full expression */
                );
            );
        );
        """)
        # Grab the result from JMP
        result = jmp.GetJSLValue("prop_string")
        if jmp.HasRunCommandErrorString:
            print("Failed to export metadata")
            print(jmp.GetRunCommandErrorString)
        elif len(result):
            print(f"Saving metadata {metatarget}...")
            with open(metatarget,"w", encoding="utf-8") as metaoutput:
                metaoutput.write("Column\tProperty\tValue\n") # Header row
                metaoutput.write(result)
        else:
            print("No metadata, moving on...")
    if not os.path.isfile(csvtarget): # Don't redo what's already done
        print(f"Saving CSV {csvtarget}...")
        jmpdoc.SaveAs(csvtarget)
    jmpdoc.Close(False,"")
    print(f"Finished with {fpth}!")

print("All done! :)")

if len(failed_files):
    print(f"\nFAILED to open the following files - try opening them directly in JMP and then rerunning this script on their parent folders:")
    for fpth in failed_files:
        print(fpth)
