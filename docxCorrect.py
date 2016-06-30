# -*- coding: utf-8 -*-
import zipfile

import tempfile
import os, os.path
import shutil
import fnmatch
import win32con, win32api
import datetime,time
import ctypes
import codecs,re
import sys, locale
import glob

class Docx(object):

	def __init__(self,docx_filename):
		self.filename = docx_filename
		with open(docx_filename,'rb') as f:
			self.zip = zipfile.ZipFile(f)
			self.xml_content = self.zip.read('word/document.xml')
			if type(self.xml_content)==type(" "):
					self.xml_content = unicode(self.xml_content,"utf-8")

	def get_xml_content(self):
		return self.xml_content


			
	def save(self, output_filename=""):
		""" Create a temp directory, expand the original docx zip.
			Write the modified xml to word/document.xml
			Zip it up as the new docx
		"""
		if output_filename=="":
			output_filename = self.filename
		tmp_dir = tempfile.mkdtemp()

		self.zip = zipfile.ZipFile(self.filename)
		self.zip.extractall(tmp_dir)

		with codecs.open(os.path.join(tmp_dir,'word/document.xml'), 'w', encoding="utf-8") as f:
			f.write(self.xml_content)

		# Get a list of all the files in the original docx zipfile
		filenames = self.zip.namelist()
		# Now, create the new zip file and add all the filex into the archive
		zip_copy_filename = output_filename
		with zipfile.ZipFile(zip_copy_filename, "w",zipfile.ZIP_DEFLATED) as docx:
			for filename in filenames:
				docx.write(os.path.join(tmp_dir,filename), filename)

		# Clean up the temp dir
		shutil.rmtree(tmp_dir)

if __name__ == "__main__":
	import logging

	sys.stdout = codecs.getwriter('utf-8')(sys.stdout);
	path = os.getcwd()
	correctLast = (len(sys.argv)>2) and (sys.argv[2]=="re") 
	undoAll = (len(sys.argv)>2) and (sys.argv[2]=="undo") 
	logOnly = (len(sys.argv)>3) and (sys.argv[3]=="log")
	verbose = (len(sys.argv)>3) and (sys.argv[3]=="verbose")
	if len(sys.argv)>1:
		if sys.argv[1].endswith(".correctdocx"):
			if sys.argv[1].endswith("re.correctdocx"):
				correctLast = True
			if sys.argv[1].endswith("undo.correctdocx"):
				undoAll = True
			if "logonly." in sys.argv[1]:
				logOnly = True
			notChangedAfterDate = time.time()
			with open(sys.argv[1],'r') as f:
				for line in f:
					label = "notChangedAfter:"
					if line.startswith(label):
						notChangedAfterDate = time.strptime(line[len(label):].strip(), "%Y-%m-%d")
					if line.startswith("verbose"):
						verbose = True
						
			path = os.path.dirname(sys.argv[1])				
		else:	
			path = sys.argv[1]
	
	logfilename = path+"\\wordlog"
	logcnt = 0
	while os.path.exists(logfilename+u".%d.txt"%logcnt):
		logcnt = logcnt + 1
	
	logging.basicConfig(level=logging.INFO, filename=logfilename+u".%d.txt"%logcnt, format="%(asctime)s - %(message)s", datefmt="%H:%M:%S", filemode='w+')
	console = logging.StreamHandler()
	console.setLevel(logging.INFO)
	formatter = logging.Formatter('%(asctime)s - %(name)s - %(message)s')
	console.setFormatter(formatter)
	logging.getLogger('').addHandler(console)
			
	isArabic = re.compile(u"(<.+>)([^<]*)[ا-ی]+([^<]*)(<.+>)")
	matches = []
	
	try:
		approot = os.path.dirname(os.path.abspath(__file__))
	except NameError:  # We are the main py2exe script, not a module
		approot = os.path.dirname(os.path.abspath(sys.argv[0]))
	
	if correctLast:
		logging.info("re...")
	if logOnly:
		logging.info("log only...")
	if undoAll:
		logging.info("undo....")
	logging.info(approot)
	logging.info(path)
	for filelineno, line in enumerate(codecs.open(approot+"\\replaces.txt",'r', encoding="utf-8")):
		line = line.strip().split("@")
		if len(line)==3:
			matches.append((re.compile(line[0]),line[1],line[2].split(",")))
		
	for root, dirnames, filenames in os.walk(unicode(path)):
		for filename in fnmatch.filter(filenames, '*.docx'):
			f = os.path.join(root, filename)

			if not os.path.exists(f):
				continue
			
			attrs = win32api.GetFileAttributes(f)
			if attrs & win32con.FILE_ATTRIBUTE_HIDDEN :
				continue
				
			if verbose:
				logging.info(">"+f)

				
			backups = [a for a in os.listdir(os.path.dirname(f)) if a.endswith(".backup.docx") and a.startswith(os.path.basename(f))]
			backups.sort()
			mt = os.path.getmtime(f)
			#(mode, ino, dev, nlink, uid, gid, size, atime, mtime, ctime)
			#st = os.stat(f)
			if len(backups)>0:
				if correctLast or undoAll:
					bestf = os.path.dirname(f) +u"\\"+ unicode(backups[0])
					
					if verbose:logging.info("replacing "+bestf + " with "+ f)
					if verbose:
						logging.info(time.strftime('%m/%d/%Y', time.gmtime(os.path.getmtime(f))))
						for bk in backups:
							bkf = os.path.dirname(f) +u"\\"+ unicode(bk)
							logging.info(time.strftime('%m/%d/%Y', time.gmtime(os.path.getmtime(bkf))))

					#for bk in backups:
					#	#tmt = os.path.getmtime(bk)
					#	(mode, ino, dev, nlink, uid, gid, size, atime, tmtime, ctime) = os.stat(bestf)
					#	if tmt<bestt:
					#		bestf = bk
					#		(mode, ino, dev, nlink, uid, gid, size, atime, bestt, ctime) = os.stat(bk)
					if verbose:logging.info(time.strftime('%m/%d/%Y',notChangedAfterDate))
					if (time.gmtime(os.path.getmtime(f)) < notChangedAfterDate) and (time.gmtime(os.path.getmtime(f)) >= time.gmtime(os.path.getmtime(bestf))):
						logging.info("try to replace..."+bestf+","+f)
						if not logOnly:
							ctypes.windll.kernel32.SetFileAttributesW(bestf,32)
							os.rename(f,f+u".problem")
							os.rename(bestf,f)
							os.remove(f+u".problem")
					#time.sleep(1)
					if verbose:logging.info("done")
					
				else:
					ignore = True#False
					#for bk in backups:
					#	tmt = os.path.getmtime(bk)
					#	#sts = os.stat(bk)
					#	#logging.info(str(mt - tmt))
					#	if mt - tmt <= 600:
					#		ignore = True
					if ignore:
						logging.warning(f+u" is done.")
						continue
			if undoAll:
				continue
			try:
				lockfilename = f+".lock"
				if os.path.exists(lockfilename):
					logging.warning(f+u" is locked, ignore")
					continue
					
				if not logOnly:
					lockfile = open(lockfilename,"w")
					lockfile.write("\n\n\n")
					lockfile.close()
					del lockfile
				
				dx = Docx(f)
			except:
				logging.warning(f+u" problem openning")
				continue
			docIsArabic = isArabic.search(dx.xml_content) is not None
			if verbose:logging.info("-------------")
			logging.info(f)
			if verbose:logging.info("is it arabic:"+str(docIsArabic))
			
			
			changes = 0
			for m in matches:
				if (u"farsi" not in m[2]) or docIsArabic:
					dx.xml_content,c = m[0].subn(m[1],dx.xml_content)
					changes += c
					while (u"multiple" in m[2]) and c>0:
						dx.xml_content,c = m[0].subn(m[1],dx.xml_content)
						changes += c
			if verbose:logging.info("changes:"+str(changes))
			
			tempfilename = f+".%s.correct.docx"%(str(mt))
			try:
				if not logOnly:
					dx.save(tempfilename)
			except:
				logging.warning(f+u" problem saving file")
				if not logOnly:
					os.remove(lockfilename)
				continue
			del dx
			
			newname = f+".%s.backup.docx"%(str(mt));
			try:
				if not logOnly:
					os.remove(newname)
			except:
				pass
			
			if not logOnly:
				os.rename(f,newname)
				os.rename(tempfilename,f)
			
				time.sleep(1)
				os.remove(lockfilename)
			
				ctypes.windll.kernel32.SetFileAttributesW(newname,6)
	raw_input('Enter Any Key')
