# -*- coding: utf-8 -*-

import cmd
import os
import sys
from colorama import init as coloramainit
from colorama import Fore
from art import *
import subprocess
import re
import winreg
from elevate import elevate
import ctypes

class WROT(cmd.Cmd):
	cmdArgs = "ALL /NoCancel /Force /OSE"

	product3 = "Microsoft Office 2003"
	product7 = "Microsoft Office 2007"
	product10 = "Microsoft Office 2010"
	product13 = "Microsoft Office 2013"
	product16 = "Microsoft Office 2016"
	productC2R = "Microsoft Office Click To Run (C2R)"

	del3path = "2003\\OffScrub03.vbs"
	del7path = "2007\\OffScrub07.vbs"
	del10path = "2010\\OffScrub10.vbs"
	del13path = "2013\\OffScrub_O15msi.vbs"
	del16path = "2016\\OffScrub_O16msi.vbs"
	delC2Rpath = "C2R\\OffScrubc2r.vbs"

	cscript = "C:\\Windows\\System32\\cscript.exe /nologo"
	cscript64 = "C:\\Windows\\SysWOW64\\cscript.exe /nologo"

	productsForDelete = []

	amd64 = False

	runCmd = ""

	def isAdmin(self):
		try:
			return ctypes.windll.shell32.IsUserAnAdmin()
		except:
			return False

	def init(self):
		if self.isAdmin() == False:
			ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
			sys.exit()
		# elevate(show_console=False)
		arch = os.environ['PROCESSOR_ARCHITEW6432']
		if arch == "AMD64":
			self.amd64 = True

		if self.amd64:
			self.prompt = Fore.CYAN + "[WROT " + Fore.MAGENTA + "x64" + Fore.CYAN + "] "
		else:
			self.prompt = Fore.CYAN + "[WROT " + Fore.MAGENTA + "x32" + Fore.CYAN + "] "
		coloramainit(autoreset=True)
		# os.system("cls")
		tprint("WROT")
		print(Fore.YELLOW + "Welcome to \"Wrapper for Removing Office Tool\" by kiber.io!")
		print(Fore.GREEN + "Your random smile for now: " + art("random") + "\n")
		print(Fore.CYAN + "This script is just a wrapper for office deletion scripts from official Microsoft GitHub.\nVBScript support is required for scripts to work.\n")

		currDir = os.path.abspath(os.getcwd())

		self.del3path = currDir + "\\" + self.del3path
		self.del7path = currDir + "\\" + self.del7path
		self.del10path = currDir + "\\" + self.del10path
		self.del13path = currDir + "\\" + self.del13path
		self.del16path = currDir + "\\" + self.del16path
		self.delC2Rpath = currDir + "\\" + self.delC2Rpath

		return self

	def do_installed(self, args):
		debug = False
		if args == "":
			pass
		elif args == "debug":
			debug = True

		office3 = False
		office7 = False
		office10 = False
		office13 = False
		office16 = False
		officeC2R = False
		installedOffice = []

		software = self.getInstalledApps()
		if debug:
			self.saveSoftwareList(software)
		for app in software:
			if "Microsoft Office" in app["name"] \
			or "Microsoft Access" in app["name"] \
			or "Microsoft Excel" in app["name"] \
			or "Microsoft PowerPoint" in app["name"] \
			or "Microsoft Outlook" in app["name"] \
			or "Microsoft Word" in app["name"] \
			or "Microsoft OneNote" in app["name"]:
				if "Update for Microsoft" in app["name"] \
				or "Components" in app["name"] \
				or " MUI " in app["name"] \
				or "Proofing" in app["name"]:
					pass
				else:
					installedOffice.append(app)
			if "Click-to-Run" in app["name"]:
				officeC2R = True

		if len(installedOffice) != 0:
			print(Fore.YELLOW + "\n\tThe following MS Office installed products found:\n")
			for i in range(len(installedOffice)):
				if "2003" in installedOffice[i]["name"]:
					office3 = True
				if "2007" in installedOffice[i]["name"]:
					office7 = True
				if "2010" in installedOffice[i]["name"]:
					office10 = True
				if "2013" in installedOffice[i]["name"]:
					office13 = True
				if "2016" in installedOffice[i]["name"]:
					office16 = True
				print(Fore.GREEN + "\t\t" + str(i + 1) + ". " + installedOffice[i]["name"] + "\n")

		possibleCommands = ""
		if office3:
			possibleCommands = possibleCommands + "del3, "
		if office7:
			possibleCommands = possibleCommands + "del7, "
		if office10:
			possibleCommands = possibleCommands + "del10, "
		if office13:
			possibleCommands = possibleCommands + "del13, "
		if office16:
			possibleCommands = possibleCommands + "del16, "
		if officeC2R:
			possibleCommands = possibleCommands + "delc2r, "
		possibleCommands = re.sub(r", $", "", possibleCommands)
		if possibleCommands != "":
			print(Fore.YELLOW + "\tPossible commands for deleting products: ", end="")
			print(Fore.GREEN + possibleCommands)

	def saveSoftwareList(self, software):
		with open("software_list.txt", "a") as sfwfile:
			for app in software:
				appString = ""
				appString = "Name: " + app["name"] + "\n"
				appString = appString + "Version: " + app["version"] + "\n"
				appString = appString + "Publisher: " + app["publisher"] + "\n-----------------\n"
				sfwfile.write(appString)
		print(Fore.MAGENTA + "\tDEBUG: Software list writed to file software_list.txt")

	def parseRegistry(self, hive, flag):
	    aReg = winreg.ConnectRegistry(None, hive)
	    aKey = winreg.OpenKey(aReg, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
	                          0, winreg.KEY_READ | flag)

	    count_subkey = winreg.QueryInfoKey(aKey)[0]

	    software_list = []

	    for i in range(count_subkey):
	        software = {}
	        try:
	            asubkey_name = winreg.EnumKey(aKey, i)
	            asubkey = winreg.OpenKey(aKey, asubkey_name)
	            software['name'] = winreg.QueryValueEx(asubkey, "DisplayName")[0]

	            try:
	                software['version'] = winreg.QueryValueEx(asubkey, "DisplayVersion")[0]
	            except EnvironmentError:
	                software['version'] = 'undefined'
	            try:
	                software['publisher'] = winreg.QueryValueEx(asubkey, "Publisher")[0]
	            except EnvironmentError:
	                software['publisher'] = 'undefined'
	            software_list.append(software)
	        except EnvironmentError:
	            continue

	    return software_list

	def getInstalledApps(self):
	    return self.parseRegistry(winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_32KEY) + self.parseRegistry(winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_64KEY) + self.parseRegistry(winreg.HKEY_CURRENT_USER, 0)

	def do_force64(self, args):
		print(Fore.YELLOW + "\tForcing 64-bit mode")
		self.amd64 = True
		self.prompt = Fore.CYAN + "[WROT " + Fore.MAGENTA + "x64" + Fore.CYAN + "] "

	def do_force32(self, args):
		print(Fore.YELLOW + "\tForcing 32-bit mode")
		self.amd64 = False
		self.prompt = Fore.CYAN + "[WROT " + Fore.MAGENTA + "x32" + Fore.CYAN + "] "

	def do_exit(self, line):
		print(Fore.GREEN + "\tGoodbye!")
		sys.exit()

	def do_del3(self, args):
		self.addToDelete(self.product3, self.del3path)

	def do_del7(self, args):
		self.addToDelete(self.product7, self.del7path)

	def do_del10(self, args):
		self.addToDelete(self.product10, self.del10path)

	def do_del13(self, args):
		self.addToDelete(self.product13, self.del13path)

	def do_del16(self, args):
		self.addToDelete(self.product16, self.del16path)

	def do_delc2r(self, args):
		self.addToDelete(self.productC2R, self.delC2Rpath)

	def addToDelete(self, product, scriptpath):
		if product in self.productsForDelete:
			print(Fore.RED + "\tYou have already marked " + product + " for deletion")
			return
		if (self.runCmd != ""):
			self.runCmd = self.runCmd + " & "

		cscript = self.cscript
		if self.amd64:
			cscript = self.cscript64

		self.runCmd = self.runCmd + "" + cscript + " " + scriptpath + " " + self.cmdArgs + ""
		self.productsForDelete.append(product)
		print(Fore.CYAN + "\t" + product + " was added to the deletion list.")
		print(Fore.YELLOW + "\tAdd another product or type \"start\" to start deleting process.	")

	def do_start(self, args):
		"""start [parallel]
		Running all deletion processes in parallel.\nATTENTION! The load on the system will increase several times, and the computer may freeze"""
		parallel = False
		if args == "parallel":
			self.runCmd = self.runCmd.replace(" & ", " | ")
			parallel = True
		if (self.runCmd == ""):
			print(Fore.RED + "\tYou didn't choose anything to delete.")
			return
		self.do_show("")
		print(Fore.YELLOW + "\n\tAre you sure you want to start deleting the selected products? " + Fore.MAGENTA + "[y/N]: ", end="")
		start = input()
		if start == "y":
			print(Fore.CYAN + "\n\tDeletion has started.")
			if parallel:
				print(Fore.RED + "\tATTENTION! The script runs in PARALLEL MODE!\n\tThe load on the system will increase several times so computer may freeze")
			# print(self.runCmd)
			# os.system("\"" + self.runCmd + "\" > nul")
			subprocess.call(self.runCmd + "\"", shell=True)
		else:
			return

	def do_run(self, args):
		self.do_start(args)

	def do_EOF(self, args):
		print("exit")
		self.do_exit(args)

	def do_show(self, args):
		if len(self.productsForDelete) == 0:
			print(Fore.YELLOW + "\tYou haven't chosen anything yet")
			return
		print(Fore.YELLOW + "\tProducts selected for deletion:")
		for i in range(len(self.productsForDelete)):
			print(Fore.RED + "\t\t" + str(i + 1) + ". " + self.productsForDelete[i])

	def do_clear(self, args):
		self.productsForDelete = []
		self.runCmd = ""
		print(Fore.YELLOW + "\tThe list to delete is cleared")

	def do_cls(self, args):
		self.do_clear(args)

if __name__ == "__main__":
	WROT().init().cmdloop()