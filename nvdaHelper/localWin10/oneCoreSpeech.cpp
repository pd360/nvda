/*
C++ code to provide access to Windows OneCore voices.
This file is a part of the NVDA project.
URL: http://www.nvaccess.org/
Copyright 2016 Tyler Spivey, NV Access Limited.
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License version 2.0, as published by
    the Free Software Foundation.
    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
This license can be found at:
http://www.gnu.org/licenses/old-licenses/gpl-2.0.html
*/

#include <string>
#include <collection.h>
#include <ppltasks.h>
#include <wrl.h>
#include <robuffer.h>
#include "oneCoreSpeech.h"

using namespace std;
using namespace Platform;
using namespace Windows::Media::SpeechSynthesis;
using namespace concurrency;
using namespace Windows::Storage::Streams;
using namespace Microsoft::WRL;
using namespace Windows::Media;
using namespace Windows::Foundation::Collections;

// Undocumented interface used by Narrator to access boosted speech rates.
interface DECLSPEC_UUID("36d1caa6-9da3-4827-a6d1-53bdd2115f10")
ISpeechSynthesisUndocumented: IInspectable {
	virtual HRESULT ParseSsmlIntoText();
	virtual HRESULT __stdcall SetVoicePropertyNum(HSTRING name, long value);
};

//Globals
SpeechSynthesizer ^synth;
ocSpeech_callbackType callback;

byte* getBytes(IBuffer^ buffer) {
	// We want direct access to the buffer rather than copying it.
	// To do this, we need to get to the IBufferByteAccess interface.
	// See http://cm-bloggers.blogspot.com/2012/09/accessing-image-pixel-data-in-ccx.html
	ComPtr<IInspectable> insp = reinterpret_cast<IInspectable*>(buffer);
	ComPtr<IBufferByteAccess> bufferByteAccess;
	insp.As(&bufferByteAccess);
	byte* bytes = nullptr;
	bufferByteAccess->Buffer(&bytes);
	return bytes;
}

void __stdcall ocSpeech_initialize() {
	synth = ref new SpeechSynthesizer();
}

void __stdcall ocSpeech_setCallback(ocSpeech_callbackType fn) {
	callback = fn;
}

int __stdcall ocSpeech_speak(char16 *s) {
	String ^text = ref new String(s);
	auto markersStr = make_shared<wstring>();
	task<SpeechSynthesisStream ^>  speakTask;
	try {
		speakTask = create_task(synth->SynthesizeSsmlToStreamAsync(text));
	} catch (Platform::Exception ^e) {
		return -1;
	}
	speakTask.then([markersStr] (SpeechSynthesisStream^ speechStream) {
		Buffer^ buffer = ref new Buffer(speechStream->Size);
		IVectorView<IMediaMarker^>^ markers = speechStream->Markers;
		for (auto&& marker : markers) {
			if (markersStr->length() > 0)
				*markersStr += L"|";
			*markersStr += marker->Text->Data();
			*markersStr += L":";
			*markersStr += to_wstring(marker->Time.Duration);
		}
		auto t = create_task(speechStream->ReadAsync(buffer, speechStream->Size, Windows::Storage::Streams::InputStreamOptions::None));
		return t;
	}).then([markersStr] (IBuffer^ buffer) {
		// Data has been read from the speech stream.
		// Pass it to the callback.
		byte *bytes = getBytes(buffer);
		callback(bytes, buffer->Length, markersStr->c_str());
	}).then([] (task<void> previous) {
		// Catch any unhandled exceptions that occurred during these tasks.
		try {
			previous.get();
		} catch (Platform::Exception^ e) {
		}
	});

	return 0;
}

const wchar_t * __stdcall ocSpeech_getVoices() {
	wstring voices;
	for (int i = 0; i < synth->AllVoices->Size; ++i) {
		VoiceInformation^ info = synth->AllVoices->GetAt(i);
		voices += info->Id->Data();
		if (i != synth->AllVoices->Size - 1)
			voices += L"|";
	}
	return voices.c_str();
}

void __stdcall ocSpeech_setVoice(int index) {
	synth->Voice = synth->AllVoices->GetAt(index);
}

void __stdcall ocSpeech_setProperty(char16 *name, long val) {
	// In order to access boosted rates, we need to use an undocumented interface.
	ComPtr<IInspectable> insp = reinterpret_cast<IInspectable*>(synth);
	ComPtr<ISpeechSynthesisUndocumented> undoc;
	if (FAILED(insp.As(&undoc))) {
		return;
	}
	HSTRING h;
	if (FAILED(WindowsCreateString(name, lstrlenW(name), &h))) {
		return;
	}
	undoc->SetVoicePropertyNum(h, val);
	WindowsDeleteString(h);
}

const char16 * __stdcall ocSpeech_getCurrentVoiceLanguage() {
	return synth->Voice->Language->Data();
}
