#A part of NonVisual Desktop Access (NVDA)
#Copyright (C) 2016 Tyler Spivey, NV Access Limited
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.

import ctypes
from synthDriverHandler import SynthDriver
from logHandler import log
import config
import nvwave
import speech

MIN_RATE = -100
MAX_RATE = 100

SSML_TEMPLATE = (u'<speak version="1.0"'
	' xmlns="http://www.w3.org/2001/10/synthesis" xml:lang="{lang}">'
	'{text}'
	'</speak>')
ocSpeech_Callback = ctypes.CFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_int, ctypes.c_wchar_p)

DLL_FILE = "lib/nvdaHelperLocalWin10.dll"

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
		self._dll = ctypes.windll[DLL_FILE]
		self._dll.ocSpeech_getCurrentVoiceLanguage.restype = ctypes.c_wchar_p
		self._handle = self._dll.ocSpeech_initialize()
		self._callbackInst = ocSpeech_Callback(self._callback)
		self._dll.ocSpeech_setCallback(self._handle, self._callbackInst)
		self._dll.ocSpeech_getVoices.restype = ctypes.c_wchar_p
		#voices = self._dll.ocSpeech_getVoices(self._handle).split('|')
		self._player = nvwave.WavePlayer(1, 22050, 16, outputDevice=config.conf["speech"]["outputDevice"])
		# Initialize state.
		self._queuedSpeech = []
		self._wasCancelled = False
		self._isProcessing = False
		# Set initial rate.
		self.rate = 40

	def terminate(self):
		super(SynthDriver, self).terminate()
		# Drop the ctypes function instance for the callback,
		# as it is holding a reference to an instance method, which causes a reference cycle.
		self._dll.ocSpeech_terminate(self._handle)
		self._callbackInst = None

	def _get_rate(self):
		return self._paramToPercent(self._rate, MIN_RATE, MAX_RATE)

	def _set_rate(self, val):
		self._rate = self._percentToParam(val, MIN_RATE, MAX_RATE)
		self._dll.ocSpeech_setProperty(self._handle, u"MSTTS.SpeakRate", self._rate)

	def cancel(self):
		# Set a flag to tell the callback not to push more audio.
		self._wasCancelled = True
		log.debug("Cancelling")
		# There might be more text pending. Throw it away.
		self._queuedSpeech = []
		self._player.stop()

	def speak(self, seq):
		new = []
		for item in seq:
			if isinstance(item, basestring):
				new.append(item.replace('<', '&lt;').replace('>', '&gt;').replace('&', '&amp;'))
			elif isinstance(item, speech.IndexCommand):
				new.append('<mark name="%s"/>' % item.index)
		text = u" ".join(new)
		# OneCore speech barfs if you don't provide the language.
		lang = self._dll.ocSpeech_getCurrentVoiceLanguage(self._handle)
		text = SSML_TEMPLATE.format(lang=lang, text=text)
		if self._isProcessing:
			# We're already processing some speech, so queue this text.
			# It'll be processed once the previous text is done.
			log.debug("Already processing, queuing")
			self._queuedSpeech.append(text)
			return
		self._wasCancelled = False
		log.debug("Begin processing speech")
		self._isProcessing = True
		# ocSpeech_speak is async.
		# It will call _callback in a background thread once done.
		self._dll.ocSpeech_speak(self._handle, text)

	def _callback(self, bytes, len, markers):
		# This gets called in a background thread.
		if len > 44:
			# Strip the first 44 bytes, as this seems to be noise.
			bytes += 44
			len -= 44
		data = ctypes.string_at(bytes, len)
		if markers:
			markers = markers.split('|')
		else:
			markers = []
		last = 0

		# Push audio up to each marker so we can sync the audio with the markers.
		for marker in markers:
			if self._wasCancelled:
				break
			name, pos = marker.split(':')
			pos = int(pos)
			# pos is a time offset in 100-nanosecond units.
			# Convert this to a byte offset.
			# 10000000 100-nanosecond units in a second
			# 22050 samples per second
			# 2 bytes per sample
			# Order the equation so we don't have to do floating point.
			pos = pos * 22050 * 2 / 10000000
			# Push audio up to this marker.
			self._player.feed(data[last:pos])
			# Indicate that we've reached this marker.
			self.lastIndex = int(name)
			last = pos
		if self._wasCancelled:
			log.debug("Cancelled, stopped pushing audio")
		else:
			self._player.feed(data[last:])
			log.debug("Done pushing audio")
		self._processNext()
		return 0

	def _processNext(self):
		if self._queuedSpeech:
			text = self._queuedSpeech.pop(0)
			log.debug("Queued speech present, begin processing next")
			self._wasCancelled = False
			# ocSpeech_speak is async.
			self._dll.ocSpeech_speak(self._handle, text)
		else:
			log.debug("Done processing")
			self._isProcessing = False
