/*
Code for C dll bridge to Windows OneCore voices.
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

OcSpeech* __stdcall ocSpeech_initialize() {
	auto instance = new OcSpeech;
	instance->synth = ref new SpeechSynthesizer();
	return instance;
}

void __stdcall ocSpeech_terminate(OcSpeech* instance) {
	delete instance;
}

void __stdcall ocSpeech_setCallback(OcSpeech* instance, ocSpeech_Callback fn) {
	instance->callback = fn;
}

int __stdcall ocSpeech_speak(OcSpeech* instance, char16 *text) {
	String^ textStr = ref new String(text);
	auto markersStr = make_shared<wstring>();
	task<SpeechSynthesisStream ^>  speakTask;
	try {
		speakTask = create_task(instance->synth->SynthesizeSsmlToStreamAsync(textStr));
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
	}).then([instance, markersStr] (IBuffer^ buffer) {
		// Data has been read from the speech stream.
		// Pass it to the callback.
		byte *bytes = getBytes(buffer);
		instance->callback(bytes, buffer->Length, markersStr->c_str());
	}).then([] (task<void> previous) {
		// Catch any unhandled exceptions that occurred during these tasks.
		try {
			previous.get();
		} catch (Platform::Exception^ e) {
		}
	});

	return 0;
}

// We use BSTR because we need the string to stay around until the client is done with it
// but the caller then needs to free it.
// We can't just use malloc because the caller might be using a different CRT.
BSTR __stdcall ocSpeech_getVoices(OcSpeech* instance) {
	wstring voices;
	for (int i = 0; i < instance->synth->AllVoices->Size; ++i) {
		VoiceInformation^ info = instance->synth->AllVoices->GetAt(i);
		voices += info->Id->Data();
		voices += L":";
		voices += info->DisplayName->Data();
		if (i != instance->synth->AllVoices->Size - 1)
			voices += L"|";
	}
	return SysAllocString(voices.c_str());
}

const char16* __stdcall ocSpeech_getCurrentVoiceId(OcSpeech* instance) {
	return instance->synth->Voice->Id->Data();
}

void __stdcall ocSpeech_setVoice(OcSpeech* instance, int index) {
	instance->synth->Voice = instance->synth->AllVoices->GetAt(index);
}

void __stdcall ocSpeech_setProperty(OcSpeech* instance, char16 *name, long val) {
	// In order to access boosted rates, we need to use an undocumented interface.
	ComPtr<IInspectable> insp = reinterpret_cast<IInspectable*>(instance->synth);
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

const char16 * __stdcall ocSpeech_getCurrentVoiceLanguage(OcSpeech* instance) {
	return instance->synth->Voice->Language->Data();
}
