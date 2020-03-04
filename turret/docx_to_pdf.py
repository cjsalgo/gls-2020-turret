bashCommand = "doc2pdf files/\"CJ Salgo\".docx"
print(bashCommand)
import subprocess
process = subprocess.Popen(bashCommand.split(), stdout=subprocess.PIPE)
output, error = process.communicate()