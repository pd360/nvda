#A part of NonVisual Desktop Access (NVDA)
#Copyright (C) 2016 Tyler Spivey, NV Access Limited
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.

"""Utilities for converting NVDA speech sequences to XML.
Several synthesizers accept XML, either SSML or their own schemas.
L{SpeechXmlConverter} is the base class for conversion to XML.
You can subclass this to support specific XML schemas.
L{SsmlConverter} is an implementation for conversion to SSML.
"""

import speech
from logHandler import log

XML_ESCAPES = {
	0x3C: u"&lt;", # <
	0x3E: u"&gt;", # >
	0x26: "&amp;", # &
}

def toXmlLang(nvdaLang):
	"""Convert an NVDA language to an XML language.
	"""
	return nvdaLang.replace("_", "-")

class SpeechXmlConverter(object):
	"""Base class for conversion of NVDA speech sequences to XML.
	NVDA speech sequences are linear, but XML is hierarchical, which makes conversion challenging.
	For example, a speech sequence might change the pitch, then change the volume, then reset the pitch to default.
	In XML, resetting to default generally requires closing the tag, but that also requires closing the outer tag.
	This class transparently handles these issues, balancing the XML as appropriate.

	Subclasses implement specific XML schemas by implementing methods which handle each speech command.
	The method for a speech command should be named with the prefix "convert" followed by the command's class name.
	For example, the handler for C{IndexCommand} should be named C{convertIndexCommand}.
	These methods receive the L{speech.SpeechCommand} instance as their only argument.
	Conversion methods then call L{text}, L{setAttr}, etc. as appropriate.

	Once you have an instance of a subclass, call the L{convert} method to convert the speech sequence and return the result.
	"""

	def __init__(self, speechSequence):
		"""Constructor.
		@param speechSequence: The speech sequence to convert.
		"""
		self.speechSequence = speechSequence
		#: The converted output as it is built.
		self._out = []
		#: A stack of open tags which enclose the entire output.
		self._enclosingAllTags = []
		#: Whether any tags have changed since last time they were output.
		self._tagsChanged = False
		#: A stack of currently open tags (excluding tags which enclose the entire output).
		self._openTags = []
		#: Current tags and their attributes.
		self._tags = {}
		#: A tag (and its attributes) which should directly enclose all text henceforth.
		self._tagEnclosingText = (None, None)

	def raw(self, text):
		"""Add raw (unprocessed) output to the output.
		"""
		self._out.append(text)

	def text(self, text):
		"""Add actual text to the output.
		This will be XML escaped.
		"""
		text = unicode(text).translate(XML_ESCAPES)
		tag, attrs = self._tagEnclosingText
		if tag:
			self._openTag(tag, attrs)
		self.raw(text)
		if tag:
			self._closeTag(tag)

	def _openTag(self, tag, attrs):
		self.raw("<%s" % tag)
		for attr, val in attrs.iteritems():
			self.raw(' %s="%s"' % (attr, val))
		self.raw(">")

	def _closeTag(self, tag):
		self.raw("</%s>" % tag)

	def encloseAll(self, tag, attrs):
		"""Enclose the entire output in a tag.
		This should be called before any other output is produced; e.g. in the constructor.
		"""
		self._openTag(tag, attrs)
		self._enclosingAllTags.append(tag)

	def setAttr(self, tag, attr, val):
		"""Set an attribute for a tag.
		The tag will be added if appropriate.
		This tag will then be output with this attribute until it is removed with L{delAttr}.
		"""
		attrs = self._tags.get(tag)
		if not attrs:
			attrs = self._tags[tag] = {}
		if attrs.get(attr) != val:
			attrs[attr] = val
			self._tagsChanged = True

	def delAttr(self, tag, attr):
		"""Remove an attribute from a tag.
		If the tag has no attributes, it will be removed.
		"""
		attrs = self._tags.get(tag)
		if not attrs:
			return
		if attr not in attrs:
			return
		del attrs[attr]
		if not attrs:
			del self._tags[tag]
		self._tagsChanged = True

	def encloseTextInTag(self, tag, attrs):
		"""Directly enclose all text henceforth with this tag.
		This will occur until L{stopEnclosingTextInTag} is called.
		"""
		self._tagEnclosingText = (tag, attrs)

	def stopEnclosingTextInTag(self):
		"""Stop directly enclosing text in a tag.
		"""
		self._tagEnclosingText = (None, None)

	def _outputTags(self):
		if not self._tagsChanged:
			return
		# Just close all open tags and reopen any existing or new ones.
		for tag in reversed(self._openTags):
			self._closeTag(tag)
		del self._openTags[:]
		for tag, attrs in self._tags.iteritems():
			self._openTag(tag, attrs)
			self._openTags.append(tag)
		self._tagsChanged = False

	def convertItem(self, item):
		if isinstance(item, basestring):
			self.text(item)
		elif isinstance(item, speech.SpeechCommand):
			name = type(item).__name__
			# For example: self.convertIndexCommand
			func = getattr(self, "convert%s" % name, None)
			if not func:
				log.debugWarning("Unsupported command: %s" % item)
				return
			func(item)
		else:
			log.error("Unknown speech: %r" % item)

	def convert(self):
		"""Convert the speech sequence to XML.
		"""
		for item in self.speechSequence:
			self.convertItem(item)
			self._outputTags()
		# Close any open tags.
		for tag in reversed(self._openTags):
			self._closeTag(tag)
		for tag in self._enclosingAllTags:
			self._closeTag(tag)
		return u"".join(self._out)

class SsmlConverter(SpeechXmlConverter):
	"""Converts an NVDA speech sequence to SSML.
	"""

	def __init__(self, speechSequence, defaultLanguage):
		super(SsmlConverter, self).__init__(speechSequence)
		self.defaultLanguage = defaultLanguage
		attrs = {"version": "1.0", "xmlns": "http://www.w3.org/2001/10/synthesis",
			"xml:lang": defaultLanguage}
		self.encloseAll("speak", attrs)

	def convertIndexCommand(self, command):
		self.raw('<mark name="%d" />' % command.index)

	def convertCharacterModeCommand(self, command):
		if command.state:
			self.encloseTextInTag("say-as", {"interpret-as": "characters"})
		else:
			self.stopEnclosingTextInTag()

	def convertLangChangeCommand(self, command):
		lang = command.lang or self.defaultLanguage
		lang = toXmlLang(lang)
		self.setAttr("voice", "xml:lang", lang)

	def convertBreakCommand(self, command):
		self.raw('<break time="%dms" />' % command.time)

	def _convertProsody(self, command, attr):
		if command.multiplier == 1:
			# Returning to normal.
			self.delAttr("prosody", attr)
		else:
			self.setAttr("prosody", attr,
				"%d%%" % int(command.multiplier* 100))

	def convertPitchCommand(self, command):
		self._convertProsody(command, "pitch")
	def convertRateCommand(self, command):
		self._convertProsody(command, "rate")
	def convertVolumeCommand(self, command):
		self._convertProsody(command, "volume")

	def convertPhonemeCommand(self, command):
		self._openTag('phoneme', {"alphabet": "ipa", "ph": command.ipa})
		self.raw(command.text)
		self._closeTag("phoneme")
