# ==============================================================================
#                        Quick Start Script
# 
#       Quick start script -- get everything up and running quickly for 
#                      service desk work. DOC version.
#                   By James Britton   --   05/12/2020
# ==============================================================================
# Imports:
# ==============================================================================

from datetime import datetime as d8
import webbrowser
import subprocess
import time
import os


# ==============================================================================
#   Variables go here:
# ==============================================================================

now = d8.now()

# Below: absolute path to Firefox (or other desired browser)
firefox_path = "C:\\Program Files\\Firefox Developer Edition\\Firefox.exe"

# Registering the above path as a browser for webbrowser module.
webbrowser.register('firefox', None,webbrowser.BackgroundBrowser(firefox_path))

exitval = 0

# NOTE: the following programs will open for the user who runs the Python file
# To open things as other users (e.g., Active Directory as AD admin), we have
# to use the os.system method to leverage "runas" and enter the right password.
# See the runas_alt() function

# Define the programs to open -- if program is not in PATH, use absolute path:
programs = (["C:\\Program Files (x86)\\KeePass Password Safe 2\\KeePass.exe", 
    "C:\\Program Files\\AutoHotkey\\AutoHotkey.exe",
    "C:\\Program Files\\Remote Desktop\\msrdcw.exe",
    "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\Outlook.exe",
    'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.exe "C:\\Users\\jambri\\Desktop\\userlist.lnk"'])

# Define the pages to open, one per line:
pages = (["google.com", 
    "communities.service-now.com/ticket.do", 
    "reddit.com"])

# Define the alternate accounts that will be used with "runas"
alt_account0 = "head_office\\jabritton.adm" # First alternate account
alt_account1 = "head_office\\jabritton.adm" # Second alternate account

# Commands to be used with "runas" - alt_account0 first.
com_lines0 = (['"powershell.exe start-process mmc dsa.msc -verb runas"',
    '"powershell.exe start-process powershell -verb runas"',
    '"powershell.exe start-process C:\Temp\CmRcViewer.exe.lnk -verb runas"',
    '"powershell.exe start-process C:\Temp\WindowsTerminal.exe.lnk -verb runas"'])

# Commands to be used with "runas" - alt_account1 next.
com_lines1 = (["notepad.exe"])


# ==============================================================================
#   Functions go here:
# ==============================================================================

# Function for opening the web-browser to a url.
def openbrowser(url):
    try:
        webbrowser.get('firefox').open(url)
    except:
        print("Failed to open web browser to: " + str(url) + "!") # We have output on-screen so we know if something breaks
        exitval = 1
        print(str(exitval) + " -- error occurred. Continuing regardless.") 
        print("=========================================================================")
        print("   ")
    finally:
        print("Opening webpage: " + str(url))
        print("=========================================================================")
        print("   ")


# Function to open programs.
def openprograms(prog):
    try:
        subprocess.Popen(prog)
    except:
        print("Failed to open " + str(prog) + "!")
        exitval = 1
        print(str(exitval) + " -- error occurred. Continuing regardless.")
        print("=========================================================================")
        print("   ")
    finally:
        print("Opening program: " + str(prog))
        print("=========================================================================")
        print("   ")


# Function to open things with Runas.
def runas_alt(com_line, alt_account):
    try:
        os.system('runas /user:{} /savecred {}'.format(alt_account, com_line)) 
        print("   ")
    except:
        print("Failed to execute runas with the following parameters: {}".format(com_line))
    finally:
        print("=========================================================================")
        print("   ")


# ==============================================================================
#   Main section: 
# ==============================================================================

print("Welcome back to the grind, Lord Foxington~")
print("It's currently: " + str(now))


# ==============================================================================
#   Opening webpages: 
# ==============================================================================

# It's another day. Another paycheck. 
print("===============             Opening webpages:             ===============")

for page in pages:
    openbrowser(page)


# ==============================================================================
#   Opening programs:
# ==============================================================================

print("===============        Opening necessary programs:        ===============")

for program in programs:
    openprograms(program)
print("   ")


# ==============================================================================
#   Opening batch files or whatever is needed to open things as other users:
# ==============================================================================

print("===============        Opening Apps as Other Users        ===============")
print("   ")

for com_line in com_lines0:
    runas_alt(com_line,alt_account0)

for com_line in com_lines1:
    runas_alt(com_line,alt_account1)

# ==============================================================================
#   Should be done. Sleep for a bit, then sign off.
# ==============================================================================

time.sleep(3)

print("Finished. Go kick some tail.")
print("   ")
print("=========================================================================")
print("   ")
print("Exit value = " + str(exitval))
print("")
print("  _,-=._              /\_/|")
print("  `-.}   `=._,.-=-._.,  @ @._,")
print("     `._ _,-.   )       _,.'")
print("        `    G.m` m`m'")

time.sleep(30)
exit(exitval)
# ==============================================================================