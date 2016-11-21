#pragma once
#define export __declspec(dllexport) 

typedef int (*ocSpeech_callbackType)(byte* data, int length, const char16* markers);
extern "C" {
void export __stdcall ocSpeech_initialize();
void export __stdcall ocSpeech_setCallback(ocSpeech_callbackType fn);
export int __stdcall ocSpeech_speak(char16 *s);
export const wchar_t * __stdcall ocSpeech_getVoices(void);
export void __stdcall ocSpeech_setVoice(int i);
export void __stdcall ocSpeech_setProperty(char16 *name, long val);
export const char16 * __stdcall ocSpeech_getCurrentVoiceLanguage();
}
