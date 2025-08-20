import os
import tkinter as tk
import sys

path_to_folder = r'\\172.16.1.70\BackupShare\install\Windows\for_WSUS'

string1 = ""
string2 = ""
string3 = ""
string4 = ""
string5 = ""
string6 = ""

servicing_stack_2016 = ""
cumulative_2016 = ""
servicing_stack_2019 = ""
cumulative_2019 = ""
servicing_stack_10_1809 = ""
cumulative_10_1809 = ""
servicing_stack_10_1607 = ""
cumulative_10_1607 = ""

lines = []
servicing_stack_2016_to_txt = ""
servicing_stack_2019_to_txt = ""
servicing_stack_10_1607_to_txt = ""
servicing_stack_10_1809_to_txt = ""
servicing_stack_lines = []
servicing_stack_lines_to_txt = ""


cumulative_2016_to_txt = ""
cumulative_2019_to_txt = ""
cumulative_10_1607_to_txt = ""
cumulative_10_1809_to_txt = ""
cumulative_lines = []
cumulative_to_txt = ""

# Function to open a WSUS folder
def open_folder():
    os.startfile(path_to_folder) # This will open the folder in file explorer


    # Function to show popup message
    def show_popup():
        # Create the root window
        root = tk.Tk()
        root.title("WARNING!")
        root.geometry("400x200")

        message = "Before proceeding, ensure that the folder contains only the relevant updates.\n Path to folder: \\\\172.16.1.70\BackupShare\install\Windows\for_WSUS "
        label = tk.Label(root, text=message, wraplength=380)
        label.pack(padx=10, pady=10)

        # Create a Button to continue
        continue_button = tk.Button(root, text="Continue", command=root.quit)
        continue_button.pack(pady=10)

        # Create a Button to open the folder
        open_folder_button = tk.Button(root, text="Open WSUS folder", command=open_folder)
        open_folder_button.pack(pady=10)

        # Create a Quit button
        quit_button = tk.Button(root, text="Quit", command=sys.exit)
        quit_button.pack(pady=10)

        # Start the Tkinter event loop
        root.mainloop()


        # Call the function to show the popup
        show_popup()

        # List all files in the folder
        try:
            files = os.listdir(path_to_folder)
            for file in files:
                file = file.strip()
            # Ensure we are dealing with .msu files
            if file.endswith('.msu'):
                # print(f"Processing file: {file}")
                if "Servicing Stack Update" in file:
                    # print(f"Found 'Service Stack Update' in: {file}")
                    if "Windows Server 2016" in file:
                        servicing_stack_2016 = file
                    if "Windows Server 2019" in file:
                        servicing_stack_2019 = file
                    if "Windows 10 Version 1607" in file:
                        servicing_stack_10_1607 = file
                    if "Windows 10 Version 1809" in file:
                        servicing_stack_10_1809 = file
                if "Cumulative Update" in file:
                    # print(f"Found 'Cumulative Update' in: {file}")
                    if "Windows Server 2016" in file:
                        cumulative_2016 = file
                    if "Windows Server 2019" in file:
                        cumulative_2019 = file
                    if "Windows 10 Version 1607" in file:
                        cumulative_10_1607 = file
                    if "Windows 10 Version 1809" in file:
                        cumulative_10_1809 = file

                    # Print results
                    # print("")
                    # print(f"servicing_stack_2016 = {servicing_stack_2016}")
                    # print(f"servicing_stack_2019 = {servicing_stack_2019}")
                    # print(f"servicing_stack_10_1809 = {servicing_stack_10_1809}")
                    # print(f"servicing_stack_10_1607 = {servicing_stack_10_1607}")
                    # print(f"cumulative_2016 = {cumulative_2016}")
                    # print(f"cumulative_2019 = {cumulative_2019}")
                    # print(f"cumulative_10_1809 = {cumulative_10_1809}")
                    # print(f"cumulative_10_1607 = {cumulative_10_1607}")

        except Exception as e:
            print(f"Error: {e}")

        if servicing_stack_2016:
            string1 = "d:\\PSTools\\psexec.exe " + "\\\\71AW02,71AW03" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(servicing_stack_2016) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string2 = "d:\\PSTools\\psexec.exe " + "\\\\71AW04,71AW05" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(servicing_stack_2016) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string3 = "d:\\PSTools\\psexec.exe " + "\\\\HIST1A,EPOSRV" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(servicing_stack_2016) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string4 = "d:\\PSTools\\psexec.exe " + "\\\\WSUS,NETMON" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(servicing_stack_2016) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string5 = "d:\\PSTools\\psexec.exe " + "\\\\PI_MAG,PIMAG_IN" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(servicing_stack_2016) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string6 = "d:\\PSTools\\psexec.exe " + "\\\\SYSADV" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(servicing_stack_2016) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'

            lines = [string1, string2, string3, string4, string5, string6]
            servicing_stack_2016_to_txt = "\n".join(lines)

        if servicing_stack_2019:
            string1 = "d:\\PSTools\\psexec.exe " + "\\\\NAS_SRV,AUTOSAVE" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(servicing_stack_2016) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string2 = "d:\\PSTools\\psexec.exe " + "\\\\VIRT03" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(servicing_stack_2019) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            lines = [string1, string2]
            servicing_stack_2019_to_txt = "\n".join(lines)

        if servicing_stack_10_1607:
            string1 = "d:\\PSTools\\psexec.exe " + "\\\\71ES02,71ES03" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(servicing_stack_10_1607) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string2 = "d:\\PSTools\\psexec.exe " + "\\\\71WS01,71WS02" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(servicing_stack_10_1607) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string3 = "d:\\PSTools\\psexec.exe " + "\\\\71WS03,71WS04" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(servicing_stack_10_1607) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string4 = "d:\\PSTools\\psexec.exe " + "\\\\71WS05,71WS06" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(servicing_stack_10_1607) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string5 = "d:\\PSTools\\psexec.exe " + "\\\\TCLI01,TCLI02" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(servicing_stack_10_1607) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string6 = "d:\\PSTools\\psexec.exe " + "\\\\TCLI03,TCLI04" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(servicing_stack_10_1607) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            lines = [string1, string2, string3, string4, string5, string6]
            lines = [string1, string2, string3, string4, string5, string6]
            servicing_stack_10_1607_to_txt = "\n".join(lines)

        if servicing_stack_10_1809:
            string1 = "d:\\PSTools\\psexec.exe " + "\\\\TCLI05,TCLI06" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(servicing_stack_10_1809) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            lines = [string1]
            servicing_stack_10_1809_to_txt = "\n".join(lines)

            servicing_stack_lines = [servicing_stack_2016_to_txt, servicing_stack_2019_to_txt, servicing_stack_10_1607_to_txt, servicing_stack_10_1809_to_txt]
            servicing_stack_lines_to_txt = "\n".join(servicing_stack_lines)


        if cumulative_2016:
            string1 = "d:\\PSTools\\psexec.exe " + "\\\\71AW02,71AW03" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(cumulative_2016) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string2 = "d:\\PSTools\\psexec.exe " + "\\\\71AW04,71AW05" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(cumulative_2016) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string3 = "d:\\PSTools\\psexec.exe " + "\\\\HIST1A,EPOSRV" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(cumulative_2016) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string4 = "d:\\PSTools\\psexec.exe " + "\\\\NETMON" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(cumulative_2016) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string5 = "d:\\PSTools\\psexec.exe " + "\\\\PI_MAG,PIMAG_IN" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(cumulative_2016) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string6 = "d:\\PSTools\\psexec.exe " + "\\\\SYSADV" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(cumulative_2016) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            lines = [string1, string2, string3, string4, string5, string6]
            cumulative_2016_to_txt = "\n".join(lines)

        if cumulative_2019:
            string1 = "d:\\PSTools\\psexec.exe " + "\\\\NAS_SRV,AUTOSAVE" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(cumulative_2019) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string2 = "d:\\PSTools\\psexec.exe " + "\\\\VIRT03" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(cumulative_2019) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            lines = [string1, string2]
            cumulative_2019_to_txt = "\n".join(lines)

        if cumulative_10_1607:
            string1 = "d:\\PSTools\\psexec.exe " + "\\\\71ES02,71ES03" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(cumulative_10_1607) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string2 = "d:\\PSTools\\psexec.exe " + "\\\\71WS01,71WS02" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(cumulative_10_1607) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string3 = "d:\\PSTools\\psexec.exe " + "\\\\71WS03,71WS04" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(cumulative_10_1607) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string4 = "d:\\PSTools\\psexec.exe " + "\\\\71WS05,71WS06" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(cumulative_10_1607) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string5 = "d:\\PSTools\\psexec.exe " + "\\\\TCLI01,TCLI02" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(cumulative_10_1607) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            string6 = "d:\\PSTools\\psexec.exe " + "\\\\TCLI03,TCLI04" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(cumulative_10_1607) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            lines = [string1, string2, string3, string4, string5, string6]
            cumulative_10_1607_to_txt = "\n".join(lines)

        if cumulative_10_1809:
            string1 = "d:\\PSTools\\psexec.exe " + "\\\\TCLI05,TCLI06" + " -s wusa /install " + '"{}'.format(path_to_folder) + "\\" + '{}"'.format(cumulative_10_1809) + ' /quiet /norestart' + ' >> "%FilePath%" 2>&1'
            lines = [string1]
            cumulative_10_1809_to_txt = "\n".join(lines)

            cumulative_lines = [cumulative_2016_to_txt, cumulative_2019_to_txt, cumulative_10_1607_to_txt, cumulative_10_1809_to_txt]
            cumulative_to_txt = "\n".join(cumulative_lines)

        # print(servicing_stack_lines_to_txt)
        # print(cumulative_to_txt)


        # Get the path to the desktop
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")


        # Define the file name
        file_name = "Windows Security Update.txt"

        # Create and write to the file on Desktop
        file_path = os.path.join(desktop_path, file_name)
        with open(file_path, 'w') as file:
            file.write("{}\n{}".format(servicing_stack_lines_to_txt,cumulative_to_txt))


            output_txt_servicing_stack_file_path = os.path.join(desktop_path, "output - servicing stack.txt")
            output_txt_cumulative_file_path = os.path.join(desktop_path, "output - cumulative.txt")
            batch_file_name_servicing_stack = "Windows Security Update - servicing stack.bat"
            batch_file_name_cumulative = "Windows Security Update - cumulative.bat"
            batch_file_servicing_stack_path = os.path.join(desktop_path, batch_file_name_servicing_stack)
            batch_file_cumulative_path = os.path.join(desktop_path, batch_file_name_cumulative)

            batch_commands_servicing_stack = """@echo off
            set FilePath={}
            echo Checking Desktop path: %FilePath%
            echo. 

            :: Overwrite or create output.txt file
            echo "Starting Batch File" > "%FilePath%"
            echo.

            {}

            echo Done updating. Check "output - servicing stack.txt" on the desktop to see the log.
            pause
            """.format(output_txt_servicing_stack_file_path, servicing_stack_lines_to_txt) # echo. - promote a new line
        with open(batch_file_servicing_stack_path, 'w') as file:
            file.write(batch_commands_servicing_stack)


            batch_commands_cumulative = """@echo off
            set FilePath={}
            echo Checking Desktop path: %FilePath%
            echo. 

            :: Overwrite or create output.txt file
            echo "Starting Batch File" > "%FilePath%"
            echo.

            {}

            echo Done updating. Check "output - cumulative.txt" on the desktop to see the log.
            pause
            """.format(output_txt_cumulative_file_path, cumulative_to_txt) # echo. - promote a new line
        with open(batch_file_cumulative_path, 'w') as file:
            file.write(batch_commands_cumulative)