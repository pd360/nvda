#include <iostream>
#include <string>
#include <collection.h>
#include <list>
#include <stdio.h>
#include <ppltasks.h>
#include <windows.h>
#include <wrl.h>
#include <robuffer.h>
#include "oneCoreSpeech.h"

using namespace Platform;
using namespace Windows::Media::SpeechSynthesis;
using namespace concurrency;
using namespace Windows::Storage::Streams;
using namespace Microsoft::WRL;
using namespace Windows::Media;
using namespace Windows::Foundation::Collections;

interface DECLSPEC_UUID("36d1caa6-9da3-4827-a6d1-53bdd2115f10")
ISomething: IInspectable {
virtual HRESULT ParseSsmlIntoText();
virtual HRESULT __stdcall SetVoicePropertyNum(HSTRING name, long value);
};

//Globals
SpeechSynthesizer ^synth;
callbackType callback;

byte* getBytes(IBuffer^ buffer)
{
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

void __stdcall initialize()
{
synth = ref new SpeechSynthesizer();
}

void __stdcall set_callback(callbackType fn)
{
callback = fn;
}

int __stdcall speak(char16 *s)
{
String ^text = ref new String(s);
std::wstring* markersStr;
task<SpeechSynthesisStream ^>  speakTask;
try {
speakTask = create_task(synth->SynthesizeSsmlToStreamAsync(text));
markersStr = new std::wstring;
} catch (Platform::Exception ^e) {
return -1;
}
speakTask.then([markersStr](SpeechSynthesisStream^ speechStream) {
Buffer^ buffer = ref new Buffer(speechStream->Size);
IVectorView<IMediaMarker^>^ markers = speechStream->Markers;
for (auto&& marker : markers) {
if (markersStr->length() > 0)
*markersStr += L"|";
*markersStr += marker->Text->Data();
*markersStr += L":";
*markersStr += std::to_wstring(marker->Time.Duration);
}
auto t = create_task(speechStream->ReadAsync(buffer, speechStream->Size, Windows::Storage::Streams::InputStreamOptions::None));
return t;
}).then([markersStr](IBuffer^ b) {
// Data has been read from the speech stream.
// Pass it to the callback.
byte *ptr = getBytes(b);
callback(ptr, b->Length, markersStr->c_str());
}).then([markersStr](task<void> previous) {
// All done. Clean up.
try {
previous.get();
}
catch (Platform::Exception^ e) {
}
delete markersStr;
});

return 0;
}

std::wstring voices;
const wchar_t * __stdcall get_voices()
{
voices = L"";
for (int i=0;i < synth->AllVoices->Size; i++) {
VoiceInformation^ info = synth->AllVoices->GetAt(i);
voices += info->DisplayName->Data();
if (i != synth->AllVoices->Size - 1)
voices += L"|";
}
return voices.c_str();
}

void __stdcall set_voice(int i)
{
synth->Voice = synth->AllVoices->GetAt(i);
}

void __stdcall set_property(char16 *name, long val)
{
ComPtr<IInspectable> insp = reinterpret_cast<IInspectable*>(synth);
ComPtr<ISomething> something;
if (FAILED(insp.As(&something))) {
printf("Error\n");
return;
}
HSTRING h;
if (FAILED(WindowsCreateString(name, lstrlenW(name), &h))) {
printf("error\n");
return;
}
//wprintf(L"%s\n", WindowsGetStringRawBuffer(h, nullptr));
//Beep(2000, 50);
something->SetVoicePropertyNum(h, val);
WindowsDeleteString(h);
}

const char16 * __stdcall get_current_voice_language() {
return synth->Voice->Language->Data();
}
