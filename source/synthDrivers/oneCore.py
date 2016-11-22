#A part of NonVisual Desktop Access (NVDA)
#Copyright (C) 2016 Tyler Spivey, NV Access Limited
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.

import os
from synthDriverHandler import SynthDriver
import ctypes
import Queue
import threading
from logHandler import log
import config
import nvwave
import speech

MIN_RATE = -100
MAX_RATE = 100

dll = None
lock = threading.RLock()
queuedSpeech = []
wasCancelled = False
isProcessing = False
lastindex = None

bgQueue = Queue.Queue()

class BgThread(threading.Thread):
	def __init__(self):
		threading.Thread.__init__(self)
		self.setDaemon(True)

	def run(self):
		while True:
			log.debug("queue get")
			func, args, kwargs = bgQueue.get()
			if not func:
				break
			try:
				log.debug("run func")
				func(*args, **kwargs)
			except:
				log.error("Error running function from queue", exc_info=True)
			bgQueue.task_done()

def _bgExec(func, *args, **kwargs):
	global bgQueue
	bgQueue.put((func, args, kwargs))

dll_file = "lib/nvdaHelperLocalWin10.dll"

class SynthDriver(SynthDriver):
	name = "oneCore"
	# Translators: Description for a speech synthesizer.
	description = _("Windows OneCore voices")
	supportedSettings = (SynthDriver.RateSetting(),)

	@classmethod
	def check(cls):
		return True

	def __init__(self):
		super(SynthDriver, self).__init__()
		self.event = threading.Event()
		self.bgt = BgThread()
		self.bgt.start()
		self.load_dll()
		if not self.event.wait(4):
			raise RuntimeError("Dll load failed or took too long")
		global player
		player = nvwave.WavePlayer(1, 22050, 16, outputDevice=config.conf["speech"]["outputDevice"])
		self.rate = 40

	def load_dll(self):
		global dll
		dll = ctypes.windll[dll_file]
		dll.ocSpeech_getCurrentVoiceLanguage.restype = ctypes.c_wchar_p
		dll.ocSpeech_initialize()
		dll.ocSpeech_setCallback(callback)
		dll.ocSpeech_getVoices.restype = ctypes.c_wchar_p
		voices = dll.ocSpeech_getVoices().split('|')
		for i, v in enumerate(voices):
			print i, v
		self.event.set()

	def _get_rate(self):
		return self._paramToPercent(self._rate, MIN_RATE, MAX_RATE)

	def _set_rate(self, val):
		self._rate = self._percentToParam(val, MIN_RATE, MAX_RATE)
		_bgExec(dll.ocSpeech_setProperty, u"MSTTS.SpeakRate", self._rate)

	def cancel(self):
		global wasCancelled, queuedSpeech
		with lock:
			wasCancelled = True
			log.debug("Cancelling")
			queuedSpeech = []
		player.stop()

	def speak(self, seq):
		new = []
		for item in seq:
			if isinstance(item, basestring):
				new.append(item.replace('<', '&lt;').replace('>', '&gt;').replace('&', '&amp;'))
			elif isinstance(item, speech.IndexCommand):
				new.append('<mark name="%s"/>' % item.index)
		text = u" ".join(new)
		lang = dll.ocSpeech_getCurrentVoiceLanguage()
		xml=u"""<speak version="1.0"
xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='%s'>
%s
</speak>"""
		text = xml % (lang, text)
		global isProcessing, wasCancelled
		with lock:
			if isProcessing:
				# We're already processing some speech, so queue this text.
				# It'll be processed once the previous text is done.
				log.debug("Already processing, queuing")
				queuedSpeech.append(text)
				return
			wasCancelled = False
			log.debug("Begin processing speech")
			isProcessing = True
		_bgExec(dll.ocSpeech_speak, text)

	def _get_lastIndex(self):
		return lastindex

@ctypes.CFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_int, ctypes.c_wchar_p)
def callback(bytes, len, markers):
		global wasCancelled, isProcessing
		if len > 44:
			bytes += 44
			len -= 44
		data = ctypes.string_at(bytes, len)
		if markers:
			markers = markers.split('|')
		else:
			markers = []
		last = 0

		for marker in markers:
			if wasCancelled:
				break
			name, t = marker.split(':')
			t = int(t)
			t = int((22050.0/10000000)*t)
			_bgExec(player.feed, data[last*2:t*2])
			_bgExec(set_last, int(name))
			last = t
		if wasCancelled:
			log.debug("Cancelled, stopped feeding")
		else:
			_bgExec(player.feed, data[last*2:])
			log.debug("Done feeding")
		log.debug("Calling done")
		_bgExec(done)
		return 0

def set_last(x):
	global lastindex
	lastindex = x

def done():
		global isProcessing, wasCancelled
		log.debug("Done called")
		with lock:
			if queuedSpeech:
				text = queuedSpeech.pop(0)
				log.debug("Queued speech present, begin processing next")
				wasCancelled = False
				dll.ocSpeech_speak(text)
			else:
				log.debug("Done processing")
				isProcessing = False
