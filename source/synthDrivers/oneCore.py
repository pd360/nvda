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

additional_text = []
MIN_RATE = -100
MAX_RATE = 100
dll = None
lock = threading.RLock()
speaking = False
lastindex = None

bgQueue = Queue.Queue()

class BgThread(threading.Thread):
	def __init__(self):
		threading.Thread.__init__(self)
		self.setDaemon(True)

	def run(self):
		global isSpeaking
		while True:
			func, args, kwargs = bgQueue.get()
			if not func:
				break
			try:
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
		_bgExec(self.load_dll)
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
		global speaking, additional_text
		with lock:
			log.info("Setting speaking to False")
			speaking = False
		clear_queue(bgQueue)
		additional_text = []
		player.stop()

	def speak(self, seq):
		new = []
		lastmark = None
		for item in seq:
			if isinstance(item, basestring):
				new.append(item.replace('<', '&lt;').replace('>', '&gt;').replace('&', '&amp;'))
			elif isinstance(item, speech.IndexCommand):
				new.append('<mark name="%s"/>' % item.index)
		text = u" ".join(new)
		_bgExec(self._speak, text)

	def _speak(self, text):
		global speaking
		lang = dll.ocSpeech_getCurrentVoiceLanguage()
		xml=u"""<speak version="1.0"
xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='%s'>
%s
</speak>"""
		text = xml % (lang, text)
		with lock:
			if speaking:
				additional_text.append(text)
				return
			log.info("Setting speaking to True")
			speaking = True

			dll.ocSpeech_speak(text)

	def _get_lastIndex(self):
		return lastindex

@ctypes.CFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_int, ctypes.c_wchar_p)
def callback(bytes, len, markers):
		global speaking
		log.info("speaking: %r" % speaking)
		data = ctypes.string_at(bytes, len)
		if markers:
			markers = markers.split('|')
		else:
			markers = []
		last = 0

		for marker in markers:
			if not speaking:
				return 0
			name, t = marker.split(':')
			t = int(t)
			t = int((22050.0/10000000)*t)
			_bgExec(player.feed, data[44+(last*2):44+(t*2)])
			_bgExec(set_last, int(name))
			last = t
		if not speaking:
			return 0
		_bgExec(player.feed, data[44+last*2:])
		_bgExec(done)
		return 0

def set_last(x):
	global lastindex
	lastindex = x

def done():
		global speaking, additional_text
		with lock:
			if speaking and not additional_text:
				log.info("Speaking to False")
				speaking = False
				return
			if speaking and additional_text:
				t = additional_text.pop(0)
				log.info("Pushing more")
				_bgExec(dll.ocSpeech_speak, t)

def clear_queue(queue):
	try:
		while True:
			queue.get_nowait()
	except:
		pass
