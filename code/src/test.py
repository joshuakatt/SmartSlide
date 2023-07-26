import subprocess

ppt_path = "/Users/joshuakattapuram/Developer/AI-presentation-handler/code/src/text_test1.pptx"

command = f'''
tell application "Terminal"
    activate
    do script "open -a 'Microsoft PowerPoint' '{ppt_path}'"
end tell
'''

process = subprocess.Popen(["osascript", "-e", command], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
output, error = process.communicate()

if error:
    print("Error:", error.decode())
