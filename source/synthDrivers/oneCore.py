#A part of NonVisual Desktop Access (NVDA)
#Copyright (C) 2016 Tyler Spivey, NV Access Limited
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.

"""Synth driver for Windows OneCore voices.
"""

import os
from collections import OrderedDict
import ctypes
import _winreg
from synthDriverHandler import SynthDriver, VoiceInfo
from logHandler import log
import config
import nvwave
import speech
import speechXml

MIN_RATE = -100
MAX_RATE = 100
MIN_PITCH = -100
MAX_PITCH = 100
SAMPLES_PER_SEC = 22050
BITS_PER_SAMPLE = 16
BYTES_PER_SEC = SAMPLES_PER_SEC * (BITS_PER_SAMPLE / 8)
#: The number of 100-nanosecond units in 1 second.
HUNDRED_NS_PER_SEC = 10000000 # 1000000000 ns per sec / 100 ns
#: The number of bytes to strip from the start of the speech output.
STRIP_OUT_START_LEN = 44
ocSpeech_Callback = ctypes.CFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_int, ctypes.c_wchar_p)
DLL_FILE = "lib/nvdaHelperLocalWin10.dll"

def bstrReturn(address):
	"""Handle a BSTR returned from a ctypes function call.
	This includes freeing the memory.
	"""
	# comtypes.BSTR.from_address seems to cause a crash for some reason. Not sure why.
	# Just access the string ourselves.
	val = ctypes.wstring_at(address)
	ctypes.windll.oleaut32.SysFreeString(address)
	return val

class _OcSsmlConverter(speechXml.SsmlConverter):

	def __init__(self, speechSequence, defaultLanguage, volume):
		super(_OcSsmlConverter, self).__init__(speechSequence, defaultLanguage)
		self._volume = volume
		self.convertVolumeCommand(speech.VolumeCommand(multiplier=1))

	def convertVolumeCommand(self, command):
		# Even the base volume must be set via SSML.
		val = self._volume * command.multiplier
		self.setAttr("prosody", "volume", val)

	def convertCharacterModeCommand(self, command):
		# OneCore's character speech sounds weird and doesn't support pitch alteration.
		# Therefore, we don't use it.
		pass

class SynthDriver(SynthDriver):
	name = "oneCore"
	# Translators: Description for a speech synthesizer.
	description = _("Windows OneCore voices")
	supportedSettings = (
		SynthDriver.VoiceSetting(),
		SynthDriver.RateSetting(),
		SynthDriver.PitchSetting(),
		SynthDriver.VolumeSetting(),
	)

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
		self._dll.ocSpeech_getVoices.restype = bstrReturn
		self._dll.ocSpeech_getCurrentVoiceId.restype = ctypes.c_wchar_p
		self._player = nvwave.WavePlayer(1, SAMPLES_PER_SEC, BITS_PER_SAMPLE, outputDevice=config.conf["speech"]["outputDevice"])
		# Initialize state.
		self._queuedSpeech = []
		self._wasCancelled = False
		self._isProcessing = False
		# Set initial values for parameters that can't be queried.
		# This initialises our cache for the value.
		self.rate = 50
		self.pitch = 50

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

	def speak(self, speechSequence):
		text = _OcSsmlConverter(speechSequence, self.language, self._volume).convert()
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
		if len > STRIP_OUT_START_LEN:
			# Strip the first 44 bytes, as this seems to be noise.
			bytes += STRIP_OUT_START_LEN
			len -= STRIP_OUT_START_LEN
		data = ctypes.string_at(bytes, len)
		if markers:
			markers = markers.split('|')
		else:
			markers = []
		prevMarker = None
		prevPos = 0

		# Push audio up to each marker so we can sync the audio with the markers.
		for marker in markers:
			if self._wasCancelled:
				break
			name, pos = marker.split(':')
			pos = int(pos)
			# pos is a time offset in 100-nanosecond units.
			# Convert this to a byte offset.
			# Order the equation so we don't have to do floating point.
			pos = pos * BYTES_PER_SEC / HUNDRED_NS_PER_SEC
			# Push audio up to this marker.
			self._player.feed(data[prevPos:pos])
			# _player.feed blocks until the previous chunk of audio is complete, not the chunk we just pushed.
			# Therefore, indicate that we've reached the previous marker.
			if prevMarker:
				self.lastIndex = prevMarker
			prevMarker = int(name)
			prevPos = pos
		if self._wasCancelled:
			log.debug("Cancelled, stopped pushing audio")
		else:
			self._player.feed(data[prevPos:])
			if prevMarker:
				self.lastIndex = prevMarker
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

	def _getAvailableVoices(self, onlyValid=True):
		voices = OrderedDict()
		voicesStr = self._dll.ocSpeech_getVoices(self._handle).split('|')
		for voiceStr in voicesStr:
			id, name = voiceStr.split(":")
			if onlyValid and not self._isVoiceValid(id):
				continue
			voices[id] = VoiceInfo(id, name)
		return voices

	def _isVoiceValid(self, id):
		idParts = id.split('\\')
		rootKey = getattr(_winreg, idParts[0])
		subkey = "\\".join(idParts[1:])
		try:
			hkey = _winreg.OpenKey(rootKey, subkey)
		except WindowsError as e:
			log.debugWarning("Could not open registry key %s, %s" % (id, e))
			return False
		try:
			langDataPath = _winreg.QueryValueEx(hkey, 'langDataPath')
		except WindowsError as e:
			log.debugWarning("Could not open registry value 'langDataPath', %s" % e)
			return False
		if not langDataPath or not isinstance(langDataPath[0], basestring):
			log.debugWarning("Invalid langDataPath value")
			return False
		if not os.path.isfile(os.path.expandvars(langDataPath[0])):
			log.debugWarning("Missing language data file: %s" % langDataPath[0])
			return False
		try:
			voicePath = _winreg.QueryValueEx(hkey, 'voicePath')
		except WindowsError as e:
			log.debugWarning("Could not open registry value 'langDataPath', %s" % e)
			return False
		if not voicePath or not isinstance(voicePath[0],basestring):
			log.debugWarning("Invalid voicePath value")
			return False
		if not os.path.isfile(os.path.expandvars(voicePath[0] + '.apm')):
			log.debugWarning("Missing voice file: %s" % voicePath[0] + ".apm")
			return False
		return True

	def _get_voice(self):
		return self._dll.ocSpeech_getCurrentVoiceId(self._handle)

	def _set_voice(self, id):
		voices = self._getAvailableVoices(onlyValid=False)
		for index, voice in enumerate(voices):
			if voice == id:
				break
		else:
			raise LookupError("No such voice: %s" % id)
		self._dll.ocSpeech_setVoice(self._handle, index)

	def _get_pitch(self):
		return self._paramToPercent(self._pitch, MIN_PITCH, MAX_PITCH)

	def _set_pitch(self, val):
		self._pitch = self._percentToParam(val, MIN_PITCH, MAX_PITCH)
		self._dll.ocSpeech_setProperty(self._handle, u"MSTTS.Pitch", self._pitch)

	def _get_language(self):
		return self._dll.ocSpeech_getCurrentVoiceLanguage(self._handle)

	def _get_volume(self):
		return self._volume

	def _set_volume(self, val):
		# This is set via SSML.
		self._volume = val
